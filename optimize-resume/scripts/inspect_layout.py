#!/usr/bin/env python3
"""Measure a rendered resume PDF; never substitute metrics for visual review."""
import argparse
import json


def inspect(path, top_mm=13.4, bottom_mm=12.5):
    import fitz
    result = []
    with fitz.open(path) as pdf:
        for index, page in enumerate(pdf):
            boxes = [b[:4] for b in page.get_text('blocks') if len(b) > 4 and b[4].strip()]
            height = page.rect.height
            top = min((b[1] for b in boxes), default=0)
            bottom = max((b[3] for b in boxes), default=0)
            printable = height - (top_mm + bottom_mm) * 72 / 25.4
            used = max(0, bottom - top_mm * 72 / 25.4)
            ratio = used / printable
            flags = []
            if not boxes:
                flags.append('blank_page')
            if ratio < .85:
                flags.append('review_underfill')
            if bottom > height - bottom_mm * 72 / 25.4 + 2:
                flags.append('review_bottom_margin')
            if index > 0 and ratio < .35:
                flags.append('review_short_last_page')
            result.append({'page': index + 1, 'printable_fill': round(ratio, 3),
                           'top_text_mm': round(top * 25.4 / 72, 1),
                           'trailing_blank_mm': round((height - bottom) * 25.4 / 72, 1),
                           'flags': flags})
    return {'pages': result}


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('pdf')
    parser.add_argument('--top-mm', type=float, default=13.4)
    parser.add_argument('--bottom-mm', type=float, default=12.5)
    args = parser.parse_args()
    print(json.dumps(inspect(args.pdf, args.top_mm, args.bottom_mm), ensure_ascii=False, indent=2))
