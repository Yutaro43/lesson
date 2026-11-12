"""まりこ: Word(.docx)の文字だけを置き換えて修正版を作る。書式（太字・色・表・レイアウト）は残す。

使い方:
  python3 replace_docx.py --input 元.docx --output 元_校正済.docx --pairs pairs.json

pairs.json の形（上から順に適用。誤りは前後の文字を含めて一意にしておく）:
  [{"old": "実施させて頂きます", "new": "実施させていただきます"}, ...]

結果として各置換の件数を表示する。0件のものは「見つからず」と出るので、本人への報告に含める。
"""
import argparse
import json

from docx import Document


def iter_paragraphs(doc):
    def walk(container):
        for p in container.paragraphs:
            yield p
        for t in container.tables:
            for row in t.rows:
                for cell in row.cells:
                    yield from walk(cell)

    yield from walk(doc)
    for s in doc.sections:
        for part in (s.header, s.footer, s.first_page_header, s.first_page_footer):
            yield from walk(part)


def replace_in_paragraph(p, old, new):
    """段落内で old を new に置き換え、置換数を返す。

    1つの run に収まっていればその run だけを書き換える（書式が完全に残る）。
    複数の run にまたがる場合は、最初の run に置換後の文字を入れ、残りの run から該当部分を削る
    （書式は最初の run のものになる）。
    """
    count = 0
    while True:
        runs = p.runs
        text = "".join(r.text for r in runs)
        idx = text.find(old)
        if idx < 0 or not runs:
            return count
        end = idx + len(old)
        pos = 0
        first = True
        for r in runs:
            r_start, r_end = pos, pos + len(r.text)
            pos = r_end
            if r_end <= idx or r_start >= end:
                continue
            a = max(idx, r_start) - r_start
            b = min(end, r_end) - r_start
            r.text = r.text[:a] + (new if first else "") + r.text[b:]
            first = False
        count += 1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--output", required=True)
    ap.add_argument("--pairs", required=True)
    args = ap.parse_args()

    with open(args.pairs, encoding="utf-8") as f:
        pairs = json.load(f)

    doc = Document(args.input)
    paragraphs = list(iter_paragraphs(doc))
    for pr in pairs:
        n = sum(replace_in_paragraph(p, pr["old"], pr["new"]) for p in paragraphs)
        status = f"{n}件" if n else "見つからず"
        print(f"{status}\t{pr['old']} → {pr['new']}")
    doc.save(args.output)
    print(f"保存: {args.output}")


if __name__ == "__main__":
    main()
