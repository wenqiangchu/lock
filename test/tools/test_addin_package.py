import argparse
import zipfile
from pathlib import Path
from xml.etree import ElementTree


CUSTOM_UI_NS = "http://schemas.microsoft.com/office/2009/07/customui"
RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
CONTENT_TYPE = "application/vnd.ms-office.customUI+xml"
EXPECTED_ACTIONS = {"RibbonStartEdit", "RibbonSubmit", "RibbonAbandon"}


def local_name(tag):
    return tag.rsplit("}", 1)[-1]


def verify_package(package_path):
    with zipfile.ZipFile(package_path) as package:
        names = set(package.namelist())
        assert "xl/vbaProject.bin" in names, "VBA project part is missing"
        assert "customUI/customUI14.xml" in names, "Ribbon part is missing"

        ribbon = ElementTree.fromstring(package.read("customUI/customUI14.xml"))
        assert ribbon.tag == f"{{{CUSTOM_UI_NS}}}customUI", "Unexpected Ribbon namespace"
        buttons = [node for node in ribbon.iter() if local_name(node.tag) == "button"]
        assert len(buttons) == 3, f"expected three Ribbon buttons, found {len(buttons)}"
        actions = {button.attrib.get("onAction") for button in buttons}
        assert actions == EXPECTED_ACTIONS, f"unexpected Ribbon callbacks: {actions}"
        assert all(button.attrib.get("getEnabled") == "RibbonGetEnabled" for button in buttons)

        content_types = ElementTree.fromstring(package.read("[Content_Types].xml"))
        overrides = [
            node
            for node in content_types
            if local_name(node.tag) == "Override"
            and node.attrib.get("PartName") == "/customUI/customUI14.xml"
        ]
        assert len(overrides) == 1, "Ribbon content type override is missing or duplicated"
        assert overrides[0].attrib.get("ContentType") == CONTENT_TYPE

        relationships = ElementTree.fromstring(package.read("_rels/.rels"))
        ribbon_relationships = [
            node
            for node in relationships
            if local_name(node.tag) == "Relationship"
            and node.attrib.get("Type") == RELATIONSHIP_TYPE
        ]
        assert len(ribbon_relationships) == 1, "Ribbon relationship is missing or duplicated"
        assert ribbon_relationships[0].attrib.get("Target") == "customUI/customUI14.xml"


def main():
    parser = argparse.ArgumentParser(description="Verify the generated MDDT .xlam package structure.")
    parser.add_argument("package", type=Path)
    args = parser.parse_args()
    verify_package(args.package)
    print("Add-in package structure verified.")


if __name__ == "__main__":
    main()
