from statusdeck.cli import main


def test_sample_build_verify(tmp_path, capsys):
    tracker = tmp_path / "t.xlsx"
    assert main(["sample", str(tracker), "--tasks", "30"]) == 0
    assert main(["validate", str(tracker)]) == 0
    out = tmp_path / "out"
    assert main(["build", str(tracker), "--out", str(out), "--iterations", "500", "--verify"]) == 0
    text = capsys.readouterr().out
    assert "formula cells match" in text
    assert (out / "status_workbook.xlsx").exists() and (out / "status_deck.pptx").exists()


def test_build_reports_cycle(tmp_path, capsys):
    from statusdeck.sample import generate
    from statusdeck.tracker import write_tracker
    prog = generate(8, seed=1)
    prog.tasks[0].preds = [("T8", 0)]
    tracker = tmp_path / "bad.xlsx"
    write_tracker(prog, tracker)
    assert main(["build", str(tracker), "--out", str(tmp_path / "o")]) == 2
    assert "cycle" in capsys.readouterr().err
