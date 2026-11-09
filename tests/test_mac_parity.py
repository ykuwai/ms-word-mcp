"""Every Mac half takes the same arguments as the COM implementation it stands in for."""

import inspect

from ms_word_mcp.tools import (  # noqa: F401
    _common, documents, edit, equations, fields, formatting, links, lists, mac, notes, page, references, screen, tables,
    text, view)


def _com_impls():
    for module in (documents, edit, equations, fields, formatting, links, lists, notes, page, references, screen,
                   tables, text, view):
        for name, value in vars(module).items():
            if name.startswith("_") and name.endswith("_impl") and callable(value):
                yield name, value


def test_mac_halves_take_the_same_arguments():
    com = dict(_com_impls())
    assert _common.MAC_IMPLS, "the Mac halves did not register"
    for name, mac_impl in _common.MAC_IMPLS.items():
        assert name in com, f"{name} has no COM implementation"
        assert list(inspect.signature(mac_impl).parameters) == \
            list(inspect.signature(com[name]).parameters), name


def test_every_tool_has_a_mac_half():
    missing = sorted(set(dict(_com_impls())) - set(_common.MAC_IMPLS))
    assert missing == [], f"no Mac half yet: {missing}"
