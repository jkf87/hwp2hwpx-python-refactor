"""XML element builder utilities for HWPX format."""

from lxml import etree

# HWPX XML Namespaces
NS = {
    "ha": "http://www.hancom.co.kr/hwpml/2011/app",
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hp10": "http://www.hancom.co.kr/hwpml/2016/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hhs": "http://www.hancom.co.kr/hwpml/2011/history",
    "hm": "http://www.hancom.co.kr/hwpml/2011/master-page",
    "hpf": "http://www.hancom.co.kr/schema/2011/hpf",
    "dc": "http://purl.org/dc/elements/1.1/",
    "opf": "http://www.idpf.org/2007/opf/",
    "ooxmlchart": "http://www.hancom.co.kr/hwpml/2016/ooxmlchart",
    "hwpunitchar": "http://www.hancom.co.kr/hwpml/2016/HwpUnitChar",
    "epub": "http://www.idpf.org/2007/ops",
    "config": "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
    "ocf": "urn:oasis:names:tc:opendocument:xmlns:container",
    "hv": "http://www.hancom.co.kr/hwpml/2011/version",
}

# Common NSMAP used on most elements
COMMON_NSMAP = {k: v for k, v in NS.items() if k not in ("ocf", "hv")}


def make_tag(prefix, localname):
    """Create a namespaced tag string."""
    return f"{{{NS[prefix]}}}{localname}"


def sub(parent, prefix, localname, attrib=None, text=None):
    """Create a sub-element with namespace."""
    tag = make_tag(prefix, localname)
    elem = etree.SubElement(parent, tag, attrib=attrib or {})
    if text is not None:
        elem.text = text
    return elem


def root_element(prefix, localname, nsmap=None):
    """Create a root element with namespace map."""
    if nsmap is None:
        nsmap = COMMON_NSMAP
    tag = make_tag(prefix, localname)
    return etree.Element(tag, nsmap=nsmap)


def to_xml_bytes(element):
    """Serialize element to UTF-8 XML bytes with declaration."""
    return etree.tostring(
        element,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
