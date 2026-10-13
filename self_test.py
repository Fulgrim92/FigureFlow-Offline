"""Local, self-contained smoke test using a synthetic demonstration PDF."""
import tempfile
from pathlib import Path
from PIL import Image
from figureflow import process_files

src = Path(__file__).resolve().parent / "examples" / "Demo_Figure_SYNTHETIC.pdf"
assert src.is_file(), "Missing synthetic example PDF"
with tempfile.TemporaryDirectory(prefix="figureflow_test_") as tmp:
    converted, report = process_files([src], Path(tmp), 300, True, True, False)
    assert len(converted) == 1, converted
    result = converted[0]
    assert result.status == "Converted", result
    assert Path(result.svg).exists() and Path(result.png).exists()
    content = Path(result.svg).read_text(encoding="utf-8")
    assert "<path" in content, "No vector paths found in SVG"
    assert "data:image/png;base64" in content, "Microscopy test image not embedded"
    assert 'width="420pt"' in content, "Word-compatible physical width missing"
    with Image.open(result.png) as im:
        assert im.width >= 1500 and im.height >= 1000
        assert abs(im.info.get("dpi", (0, 0))[0] - 300) < 1
    assert Path(report).is_file()
    print("PASS: local PDF conversion, embedded images, vector paths, DPI, preview, report")