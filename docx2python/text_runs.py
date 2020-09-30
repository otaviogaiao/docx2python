#!/usr/bin/env python3
# _*_ coding: utf-8 _*_
"""Get text run formatting.

:author: Shay Hill
:created: 7/4/2019

Text runs are formatted inline in the ``trash/document.xml`` or header files. Read
those elements to extract formatting information.
"""
import re
from typing import Dict, List, Optional, Sequence, Tuple
from xml.etree import ElementTree

from .namespace import qn


def _elem_tag_str(elem: ElementTree.Element) -> str:
    """
    The text part of an elem.tag (the portion right of the colon)

    :param elem: an xml element

    create with::

        document = ElementTree.fromstring('bytes string')
        # recursively search document elements.

    **E.g., given**:

        document = ElementTree.fromstring('bytes string')
        # document.tag = '{http://schemas.openxml.../2006/main}:document'
        elem_tag_str(document)

    **E.g., returns**:

        'document'
        """
    return re.match(r"{.*}(?P<tag_name>\w+)", elem.tag).group("tag_name")


def gather_tcPr(run_element: ElementTree.Element) -> Dict[str, Optional[str]]:
    """
    Gather formatting elements for a table cell.

    :param run_element: a ``<w:tc>`` xml element

    create with::

        document = ElementTree.fromstring('bytes string')
        # recursively search document for <w:r> elements.

    :return: Style names ('b/', 'sz', etc.) mapped to values.

    To keep things more homogeneous, I've given tags list b/ (bold) a value of None,
    even though they don't take a value in xml.

    Each element of rPr will be either present (returned tag: None) or have a value
    (returned tag: val).

    **E.g., given**::

        <w:tc>
            <w:tcPr>
                <w:tcW w:w="2880" w:type="dxa"/>
                <w:gridSpan w:val="2"/>
            </w:tcPr>
            <w:p>
                <w:r>
                    <w:t>AAA</w:t>
                </w:r>
            </w:p>
        </w:tc>

    **E.g., returns**::

        {
            "gridSpan": "2"
        }
    """
    try:
        tcPr = run_element.find(qn("w:tcPr"))
        return {_elem_tag_str(x): x.attrib.get(qn("w:val"), None) for x in tcPr}
    except TypeError:
        # no formatting for run
        return {}

# noinspection PyPep8Naming


def get_table_cell_style(run_element: ElementTree.Element) -> List[Tuple[str, str]]:
    """Select only tcPr tags you'd like to implement.

    :param run_element: a ``<w:tc>`` xml element

    create with::

        document = ElementTree.fromstring('bytes string')
        # recursively search document for <w:tc> elements.

    :return: ``[(tcPr, val), (tcPr, val) ...]``

    Also see docstring for ``gather_tcPr``
    """
    tcPr2val = gather_tcPr(run_element)
    style = []
    for tag, val in sorted(tcPr2val.items()):
        if tag in {"gridSpan"}:
            style.append((tag, f'val="{val}"'))

    return style

# noinspection PyPep8Naming


def gather_pPr(run_element: ElementTree.Element) -> Dict[str, Optional[str]]:
    """
    Gather formatting elements for a text paragraph.

    :param run_element: a ``<w:p>`` xml element

    create with::

        document = ElementTree.fromstring('bytes string')
        # recursively search document for <w:p> elements.

    :return: Style names ('b/', 'sz', etc.) mapped to values.

    To keep things more homogeneous, I've given tags list b/ (bold) a value of None,
    even though they don't take a value in xml.

    Each element of pPr will be either present (returned tag: None) or have a value
    (returned tag: val).

    **E.g., given**::

     <w:p>
        <w:pPr>
        <w:pStyle> w:val="NormalWeb"/>
        <w:spacing w:before="120" w:after="120"/>
        </w:pPr>
        <w:r>
            <w:t xml"space="preserve">I feel that there is much to be said for the Celtic belief that the souls of those whom we have lost are held captive in some inferior being...</w:t>
        </w:r>
    </w:p>

    **E.g., returns**::

        {
            "rFonts": True,
            "spacing": 120,
            "u": "single",
            "i": None,
            "sz": "32",
            "color": "red",
            "szCs": "32",
            "s": None,
        }
    """
    try:
        pPr = run_element.find(qn("w:pPr"))
        elements = {}
        for x in pPr:
            tag = _elem_tag_str(x)
            print(tag)
            if tag == "spacing":
                elements[tag] = {'before': x.attrib.get(
                    qn("w:before"), None), 'after': x.attrib.get(qn("w:after"), None),
                    'line': x.attrib.get(qn("w:line"), None), 'lineRule': x.attrib.get(qn("w:lineRule"), None),
                    'beforeAutospacing': x.attrib.get(qn("w:beforeAutospacing"), None),
                    'afterAutospacing': x.attrib.get(qn("w:afterAutospacing"), None)}
            elif tag in {"jc"}:
                elements[tag] = x.attrib.get(qn("w:val"), None)
            elif tag == "ind":
                elements[tag] = {
                    'left': x.attrib.get(qn("w:left"), None),
                    'right': x.attrib.get(qn("w:right"), None),
                    'hanging': x.attrib.get(qn("w:hanging"), None),
                    'firstLine': x.attrib.get(qn("w:firstLine"), None),
                }
        return elements
    except TypeError:
        # no formatting for run
        return {}


# noinspection PyPep8Naming
def get_paragraph_style(run_element: ElementTree.Element) -> List[Tuple[str, str]]:
    """Select only rPr2 tags you'd like to implement.

    :param run_element: a ``<w:p>`` xml element

    create with::

        document = ElementTree.fromstring('bytes string')
        # recursively search document for <w:p> elements.

    :return: ``[(pPr, val), (pPr, val) ...]``

    Tuples are always returned in order:

    Also see docstring for ``gather_pPr``
    """
    pPr2val = gather_pPr(run_element)
    style = []
    spacings = []
    indentations = []
    for tag, val in sorted(pPr2val.items()):
        if tag == "spacing":
            for key, value in val.items():
                spacings.append(f'{key}="{value}"')
        elif tag == "ind":
            for key, value in val.items():
                indentations.append(f'{key}="{value}"')
        elif tag in {"jc"}:
            style.append(("justification", f'val="{val}"'))

    if indentations:
        style = [("indentation", " ".join(sorted(indentations)))] + style
    if spacings:
        style = [("spacing", " ".join(sorted(spacings)))] + style

    return style

# noinspection PyPep8Naming


def gather_rPr(run_element: ElementTree.Element) -> Dict[str, Optional[str]]:
    """
    Gather formatting elements for a text run.

    :param run_element: a ``<w:r>`` xml element

    create with::

        document = ElementTree.fromstring('bytes string')
        # recursively search document for <w:r> elements.

    :return: Style names ('b/', 'sz', etc.) mapped to values.

    To keep things more homogeneous, I've given tags list b/ (bold) a value of None,
    even though they don't take a value in xml.

    Each element of rPr will be either present (returned tag: None) or have a value
    (returned tag: val).

    **E.g., given**::

         <w:r w:rsidRPr="000E1B98">
            <w:rPr>
                <w:rFonts w:ascii="Arial"/>
                <w:b/>
                <w:sz w:val="32"/>
                <w:szCs w:val="32"/>
                <w:u w:val="single"/>
                <w:strike w:val="true"/>
            </w:rPr>
            <w:t>text styled  with rPr
            </w:t>
        </w:r>

    **E.g., returns**::

        {
            "rFonts": True,
            "b": None,
            "u": "single",
            "i": None,
            "sz": "32",
            "color": "red",
            "szCs": "32",
            "s": None,
        }
    """
    try:
        rPr = run_element.find(qn("w:rPr"))
        return {_elem_tag_str(x): x.attrib.get(qn("w:val"), None) for x in rPr}
    except TypeError:
        # no formatting for run
        return {}


# noinspection PyPep8Naming
def get_run_style(run_element: ElementTree.Element) -> List[Tuple[str, str]]:
    """Select only rPr2 tags you'd like to implement.

    :param run_element: a ``<w:r>`` xml element

    create with::

        document = ElementTree.fromstring('bytes string')
        # recursively search document for <w:r> elements.

    :return: ``[(rPr, val), (rPr, val) ...]``

    Tuples are always returned in order:

    ``"font"`` first then any other styles in alphabetical order.

    Also see docstring for ``gather_rPr``
    """
    rPr2val = gather_rPr(run_element)
    style = []
    font_styles = []
    for tag, val in sorted(rPr2val.items()):
        if tag in {"strike", "s"}:
            style.append(("s", ""))
        elif tag in {"b", "i", "u"}:
            style.append((tag, ""))
        elif tag == "sz":
            font_styles.append('size="{}"'.format(val))
        elif tag == "color":
            font_styles.append('color="{}"'.format(val))

    if font_styles:
        style = [("font", " ".join(sorted(font_styles)))] + style
    return style


def style_open(style: Sequence[Tuple[str, str]]) -> str:
    """
    HTML tags to open a style.

    >>> style = [
    ...     ("font", 'color="red" size="32"'),
    ...     ("b", ""),
    ...     ("i", ""),
    ...     ("u", ""),
    ... ]
    >>> style_open(style)
    '<font color="red" size="32"><b><i><u>'
    """
    text = [" ".join(x for x in y if x) for y in style]
    return "".join("<{}>".format(x) for x in text)


def style_close(style: List[Tuple[str, str]]) -> str:
    """
    HTML tags to close a style.

    >>> style = [
    ...     ("font", 'color="red" size="32"'),
    ...     ("b", ""),
    ...     ("i", ""),
    ...     ("u", ""),
    ... ]
    >>> style_close(style)
    '</u></i></b></font>'

    Tags will always be in reverse (of open) order, so open - close will look like::

        <b><i><u>text</u></i></b>
    """
    return "".join("</{}>".format(x) for x, _ in reversed(style))
