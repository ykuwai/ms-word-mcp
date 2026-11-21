"""A style's language: read it, set it, turn checking off."""

import pytest
from pydantic import ValidationError

from ms_word_mcp.tools import formatting
from ms_word_mcp.tools.checks import LANGUAGES


def test_a_style_language_in_word_set_languages_names():
    assert formatting.style_language(2057, 1041, False) == {
        "language": "english_uk", "east_asian_language": "japanese"}
    assert formatting.style_language(2057, 1041, True)["language"] == "no_proofing"
    assert "custom locale" in formatting.style_language(0x0C00, 1041, False)["language"]


def test_a_language_that_reads_back_otherwise_says_so():
    described = {"language": "english_us", "east_asian_language": "japanese"}
    assert formatting.language_warnings("english_us", "japanese", described) == []
    assert formatting.language_warnings("english_uk", None, described) == [
        "language reads back as 'english_us', not 'english_uk'."]


def test_the_languages_are_word_set_languages_and_count_as_something_to_set():
    assert formatting.SetStyleInput(style="Quote", language="english_uk").language == "english_uk"
    assert formatting.SetStyleInput(style="Quote", language="no_proofing").language == "no_proofing"
    with pytest.raises(ValidationError):
        formatting.SetStyleInput(style="Quote", language="klingon")
    with pytest.raises(ValidationError, match="no_proofing goes in `language`"):
        formatting.SetStyleInput(style="Quote", east_asian_language="no_proofing")
    assert set(formatting.SetStyleInput.model_fields["language"].annotation.__args__[0].__args__) \
        == set(LANGUAGES)
