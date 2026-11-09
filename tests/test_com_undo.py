import pytest
"""One undo step per writing call, and an honest note when Word cannot give one."""

from types import SimpleNamespace

from ms_word_mcp.backend import com
from ms_word_mcp.backend.worker import _Job


class FakeRecord:
    def __init__(self):
        self.IsRecordingCustomRecord = False
        self.started = []

    def StartCustomRecord(self, name):
        self.started.append(name)
        self.IsRecordingCustomRecord = True

    def EndCustomRecord(self):
        self.IsRecordingCustomRecord = False


def make(active_name, target_name):
    word = com.ComWord()
    record = FakeRecord()
    word.app = SimpleNamespace(Name="Word", UndoRecord=record,
                               ActiveDocument=SimpleNamespace(FullName=active_name))
    word.get_app = lambda allow_launch=False: word.app
    word.document = lambda: SimpleNamespace(FullName=target_name)
    return word, record


def run(word, func, undo="Insert paragraphs"):
    return word.run_call(_Job(func, (), {}, False, None, undo))


def test_a_call_on_the_active_document_is_one_undo_step():
    word, record = make("C:/a.docx", "C:/a.docx")
    out = run(word, lambda: {"success": True})
    assert record.started == ["Insert paragraphs"]
    assert not record.IsRecordingCustomRecord
    assert "undo" not in out


def test_a_call_on_another_document_says_it_is_not_one_step():
    word, record = make("C:/b.docx", "C:/a.docx")
    out = run(word, lambda: {"success": True})
    assert record.started == []
    assert out["undo"] == com.UNGROUPED_NOTE


def test_a_record_word_ended_early_is_reported():
    word, record = make("C:/a.docx", "C:/a.docx")

    def write():
        record.IsRecordingCustomRecord = False  # Word ended it
        return {"success": True}

    assert run(word, write)["undo"] == com.UNGROUPED_NOTE


def test_reads_and_failures_carry_no_note():
    word, record = make("C:/b.docx", "C:/a.docx")
    assert "undo" not in run(word, lambda: {"success": True}, undo=None)
    assert run(word, lambda: {"error": "x"}) == {"error": "x"}


def test_a_target_that_was_closed_is_an_error_not_the_active_document():
    word = com.ComWord()
    word.app = SimpleNamespace(Documents=SimpleNamespace(Count=1),
                               ActiveDocument=SimpleNamespace(FullName="C:\\other.docx"))
    word.get_app = lambda allow_launch=False: word.app
    word.open_documents = lambda: []
    word.target_full_name = "C:\\gone.docx"
    with pytest.raises(ValueError):
        com.ComWord.document(word)
    assert word.target_full_name == "C:\\gone.docx"
