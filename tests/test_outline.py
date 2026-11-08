"""Placing Word's heading list on its paragraphs."""

from ms_word_mcp.tools import _outline as o


def _word(paragraphs):
    """read_text and read_level over Word's real paragraphs: (text, level) each."""
    reads = []

    def read_text(i):
        reads.append(i)
        return paragraphs[i - 1][0] if 1 <= i <= len(paragraphs) else None

    def read_level(i):
        return paragraphs[i - 1][1]

    return read_text, read_level, reads


def test_the_split_numbers_a_plain_document():
    whole = "Intro\rtext\rPart\rmore\r"
    assert o.split_texts(whole) == ["Intro", "text", "Part", "more"]


def test_cells_and_row_ends_are_paragraphs():
    whole = "a\r\x07b\r\x07\r\x07after\r"
    assert o.split_texts(whole) == ["a", "b", "", "after"]


def test_items_give_level_and_text():
    assert o.parse_item("    Deep\theading") == (3, "Deep heading")
    assert o.parse_item("Top") == (1, "Top")


def test_headings_are_placed_and_each_read_back():
    paragraphs = [("Intro", 1), ("text", None), ("Part", 2), ("more", None)]
    read_text, read_level, reads = _word(paragraphs)
    found = o.locate_headings(["Intro", "  Part"], o.split_texts("Intro\rtext\rPart\rmore\r"),
                              read_text, read_level)
    assert found == [1, 3]
    assert reads == [1, 3]


def test_a_contents_line_with_the_same_words_is_passed_over():
    paragraphs = [("Intro\t1", None), ("Intro", 1), ("text", None)]
    read_text, read_level, _ = _word(paragraphs)
    whole = "Intro\t1\rIntro\rtext\r"
    assert o.locate_headings(["Intro"], o.split_texts(whole), read_text, read_level) == [2]


def test_an_equation_makes_the_split_run_ahead_and_the_heading_is_found_behind():
    # Word: 4 paragraphs. The equation's text holds marks of its own, so the
    # split has 7 pieces and puts "Next" at 6.
    paragraphs = [("Intro", 1), ("E=mc2", None), ("Next", 1), ("after", None)]
    whole = "Intro\rE\r=\rm\rc2\rNext\rafter\r"
    read_text, read_level, _ = _word(paragraphs)
    assert o.locate_headings(["Intro", "Next"], o.split_texts(whole), read_text,
                             read_level) == [1, 3]


def test_a_shortened_heading_still_matches():
    long = "A very long heading " * 10
    paragraphs = [(long.strip(), 2)]
    read_text, read_level, _ = _word(paragraphs)
    item = "  " + long[:100]
    assert o.locate_headings([item], o.split_texts(long + "\r"), read_text,
                             read_level) == [1]


def test_a_heading_that_cannot_be_found_gives_up():
    paragraphs = [("Intro", 1)]
    read_text, read_level, _ = _word(paragraphs)
    assert o.locate_headings(["Missing"], o.split_texts("Intro\r"), read_text,
                             read_level) is None


def test_a_heading_at_the_wrong_level_is_not_taken():
    paragraphs = [("Intro", 2)]
    read_text, read_level, _ = _word(paragraphs)
    assert o.locate_headings(["Intro"], o.split_texts("Intro\r"), read_text,
                             read_level) is None
