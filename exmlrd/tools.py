from typing import Dict, Optional


def del_namespace(
    tag: str,
    namespace: str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
):
    return str(tag).replace("{%s}" % namespace, "")


def set_classattr(obj, *, key="", value="", container: Optional[Dict] = None) -> None:
    """Set class attribute for an object.

    Args:
        obj (object): Object to set class attribute for.
        key (str, optional): Attribute name. Defaults to "".
        value (str, optional): Attribute value. Defaults to "".
        container (dict, optional): Dictionary of attributes to set. Defaults to None.

    Returns:
    None
    """
    if container is not None:
        for k, v in container.items():
            if hasattr(obj, k):
                setattr(obj, k, v)
    else:
        if hasattr(obj, key):
            setattr(obj, key, value)
