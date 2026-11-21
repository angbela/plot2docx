import hashlib
import io

from docx import Document
from PIL import Image
from pypdf import PdfWriter
from pypdf.generic import (
    ArrayObject,
    ContentStream,
    DecodedStreamObject,
    DictionaryObject,
    IndirectObject,
    NameObject,
    NumberObject,
)


PLOT_TAG_PREFIX = "Plot to Docx figure "


def _document_images(docx_path):
    document = Document(docx_path)
    images = {}
    for shape in document.inline_shapes:
        tag = shape._inline.docPr.get("descr", "")
        if not tag.startswith(PLOT_TAG_PREFIX):
            continue
        if tag in images:
            raise ValueError(f"Duplicate plot identifier in Word document: {tag}")
        relationship_id = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
        images[tag] = document.part.related_parts[relationship_id].blob
    if not images:
        raise ValueError("The Word document contains no identifiable plot images.")
    return images


def _figure_locations(writer, expected_tags):
    root = writer.root_object.get("/StructTreeRoot")
    if root is None:
        raise ValueError("Word did not produce the PDF structure tags needed to identify plots.")
    locations = {}
    stack = [root]
    while stack:
        item = stack.pop()
        item = item.get_object()
        if isinstance(item, ArrayObject):
            stack.extend(item)
        elif isinstance(item, DictionaryObject):
            tag = item.get("/Alt")
            if tag in expected_tags:
                page = item.get("/Pg")
                children = item.get("/K")
                if isinstance(children, ArrayObject) and len(children) == 1:
                    children = children[0]
                if page is None or not isinstance(children, NumberObject):
                    raise ValueError(f"Unsupported PDF structure for plot: {tag}")
                page_ref = page.indirect_reference
                if page_ref is None or tag in locations:
                    raise ValueError(f"Ambiguous PDF structure for plot: {tag}")
                locations[tag] = (page_ref.idnum, int(children))
            elif "/K" in item:
                stack.append(item["/K"])
    missing = expected_tags - locations.keys()
    if missing:
        raise ValueError(f"Word PDF is missing {len(missing)} plot identifier(s).")
    if len(set(locations.values())) != len(locations):
        raise ValueError("Multiple plots share the same PDF structure tag.")
    return {location: tag for tag, location in locations.items()}


def _image_references(writer, locations):
    references = {}
    for page in writer.pages:
        page_id = page.indirect_reference.idnum
        xobjects = page["/Resources"].get("/XObject", DictionaryObject()).get_object()
        active = []
        for operands, operator in ContentStream(page.get_contents(), writer).operations:
            if operator == b"BDC":
                properties = operands[1].get_object()
                mcid = properties.get("/MCID") if isinstance(properties, DictionaryObject) else None
                active.append(mcid)
            elif operator == b"BMC":
                active.append(None)
            elif operator == b"EMC":
                if not active:
                    raise ValueError("Invalid marked content in Word PDF.")
                active.pop()
            elif operator == b"Do":
                tags = [locations[(page_id, int(mcid))] for mcid in active
                        if mcid is not None and (page_id, int(mcid)) in locations]
                if not tags:
                    continue
                if len(tags) != 1 or tags[0] in references:
                    raise ValueError("Ambiguous plot image in Word PDF.")
                reference = xobjects.raw_get(operands[0])
                if not isinstance(reference, IndirectObject) or reference.get_object().get("/Subtype") != "/Image":
                    raise ValueError(f"Unsupported PDF image for plot: {tags[0]}")
                references[tags[0]] = reference
    missing = set(locations.values()) - references.keys()
    if missing:
        raise ValueError(f"Could not locate {len(missing)} plot image(s) in Word PDF.")
    return references


def _image_stream(image, color_space):
    stream = DecodedStreamObject()
    stream.update({
        NameObject("/Type"): NameObject("/XObject"),
        NameObject("/Subtype"): NameObject("/Image"),
        NameObject("/Width"): NumberObject(image.width),
        NameObject("/Height"): NumberObject(image.height),
        NameObject("/ColorSpace"): NameObject(color_space),
        NameObject("/BitsPerComponent"): NumberObject(8),
    })
    stream.set_data(image.tobytes())
    return stream.flate_encode()


def restore_pdf_images(docx_path, input_pdf, output_pdf):
    images = _document_images(docx_path)
    writer = PdfWriter(clone_from=input_pdf)
    locations = _figure_locations(writer, set(images))
    references = _image_references(writer, locations)

    replacements = {}
    for tag, reference in references.items():
        digest = hashlib.sha256(images[tag]).digest()
        if reference.idnum in replacements:
            if replacements[reference.idnum] != digest:
                raise ValueError("Word merged different plot images in the PDF; cannot restore them safely.")
            continue
        replacements[reference.idnum] = digest
        with Image.open(io.BytesIO(images[tag])) as image:
            rgba = image.convert("RGBA")
            rgb = rgba.convert("RGB")
            stream = _image_stream(rgb, "/DeviceRGB")
            alpha = rgba.getchannel("A")
            if alpha.getextrema() != (255, 255):
                mask = _image_stream(alpha, "/DeviceGray")
                stream[NameObject("/SMask")] = writer._add_object(mask)
        # Replace the referenced object so Word's original page geometry remains intact.
        stream.indirect_reference = reference
        writer._objects[reference.idnum - 1] = stream

    with open(output_pdf, "wb") as output:
        writer.write(output)
