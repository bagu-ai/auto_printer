import json

import pytest

from label_printer.core.job_journal import JobJournal, JournalError


def read_events(path):
    return [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines()]


def test_journal_flushes_intent_before_return_and_uses_unique_batch(tmp_path):
    original = {"template": "original.btw", "jobs": [{"row": 1, "values": {"id": "A"}}]}
    journal = JobJournal(tmp_path, original)
    original["jobs"][0]["values"]["id"] = "changed"
    journal.append("row_start", row=1)
    events = read_events(journal.path)
    assert [e["event"] for e in events] == ["batch_start", "row_start"]
    assert events[0]["snapshot"]["jobs"][0]["values"]["id"] == "A"
    assert [e["sequence"] for e in events] == [1, 2]
    other = JobJournal(tmp_path, {})
    assert other.batch_id != journal.batch_id
    other.close()
    journal.close()


def test_summary_preserves_start_snapshot_and_statuses(tmp_path):
    journal = JobJournal(tmp_path, {"template": "initial.btw", "printer": "initial printer"})
    journal.save_summary(False, "unknown", [{"row": 1, "status": "unknown"}])
    summary = json.loads(journal.summary_path.read_text(encoding="utf-8"))
    assert summary["template"] == "initial.btw"
    assert summary["rows"][0]["status"] == "unknown"
    assert summary["batch_id"] == journal.batch_id
    journal.close()


def test_critical_record_requires_fsync_and_closed_journal_fails(tmp_path, monkeypatch):
    journal = JobJournal(tmp_path, {})
    def fail_sync(fd):
        raise OSError("simulated disk full")
    monkeypatch.setattr("label_printer.core.job_journal.os.fsync", fail_sync)
    with pytest.raises(JournalError, match="日志写入失败"):
        journal.append("row_start", row=1)
    journal.close()
    with pytest.raises(JournalError, match="已经关闭"):
        journal.append("row_start", row=2)


def test_failed_initial_log_cannot_return_a_usable_journal(tmp_path, monkeypatch):
    def fail_sync(fd):
        raise OSError("simulated disk full")
    monkeypatch.setattr("label_printer.core.job_journal.os.fsync", fail_sync)
    with pytest.raises(JournalError):
        JobJournal(tmp_path, {})
