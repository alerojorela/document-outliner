# -*- coding: utf-8 -*-
# by Alejandro Rojo Gualix 2022-02 ...
__author__ = 'Alejandro Rojo Gualix'
"""
pip install python-docx freeplane-io
"""

import sys
import argparse
from pathlib import Path

import re
import freeplane
# python-docx
from docx import Document
from docx.enum.text import WD_COLOR_INDEX

"""
SR. No.	Colour Name In WD_COLOR_INDEX	Colour Description
1.	AUTO	Default or No Colour
2.	BLACK	Black Colour
3.	BLUE	Blue Colour
4.	BRIGHT_GREEN	Green Colour
5.	DARK_BLUE	Dark Blue Colour
6.	DARK_RED	Dark Red Colour
7.	DARK_YELLOW	Dark Yellow Colour
8.	GRAY_25	Light Gray Colour
9.	GRAY_50	Dark Gray Colour
10.	GREEN	Dark Green Colour
11.	PINK	Magenta Colour
12.	RED	Red Colour
13.	TEAL	Dark Cyan Colour
14.	TURQUOISE	Cyan Colour
15.	VIOLET	Dark Magenta Colour
16.	WHITE	White Colour
17.	YELLOW	Yellow Colour
"""
from summarizer import *

extensions = ['*.odt', '*.docx']
extensions = ['*.docx']


class Outliner:
    def __init__(self, file: Path):
        self.file = file
        # look for title?
        self.name = file.stem
        self.document = Document(file)
        self._output_document = Document()
        self._output_midmapping = None

    def extract_marks(self, doc_file: Path = None, freeplane_file: Path = None, summarizer=None,
                      select_bold=True, select_underline=False, highlighted_color=None):
        sample = ''
        for paragraph in self.document.paragraphs:
            sample += paragraph.text + '\n'
            if len(sample) > 100:
                break
        if not sample:
            return

        def text_filter(run):
            # print(run.font.highlight_color, run.text)
            return (select_bold and run.bold) or \
                (select_underline and run.underline) or \
                (highlighted_color and run.font.highlight_color == highlighted_color)

        if freeplane_file:
            self._output_midmapping = freeplane.Mindmap()
            self._output_midmapping.rootnode.plaintext = self.name
            parents_stack = [self._output_midmapping.rootnode]

        def append_summary(text):
            result = summarize(text, summarizer, max_length=min(120, int(len(text) * 0.3)))
            summary = result[0]['summary_text']
            # print('summary', summary)

            if doc_file:
                self._output_document.add_paragraph(summary)
            if freeplane_file:  # Freeplane
                node = parents_stack[-1].add_child(summary)
                node._node.attrib["STYLE"] = 'fork'

        styles = set()
        section_text = []
        for paragraph in self.document.paragraphs:
            style = paragraph.style.name
            styles.add(style)

            if style.startswith("Title") or style.startswith("Heading"):
                if summarizer and section_text:
                    compiled_text = '\n'.join(section_text)
                    append_summary(compiled_text)
                    section_text = []

                if style.startswith("Title"):
                    level = 1  # TODO 0
                elif style.startswith("Heading"):
                    m = re.search(r'(\d+)$', style)
                    level = int(m.group(0))

                if doc_file:
                    self._output_document.add_heading(paragraph.text, level)

                if freeplane_file:
                    # Freeplane
                    # print('\t', level, len(parents_stack) - 1)
                    if level > len(parents_stack) - 1:  # rise level
                        while level > len(parents_stack):
                            parents_stack.append(parents_stack[-1].add_child('<missing branch>'))
                    else:  # drop: same or lower the level
                        while len(parents_stack) - level > 0:
                            parents_stack.pop()
                    node = parents_stack[-1].add_child(paragraph.text)
                    node._node.attrib["STYLE"] = 'bubble'
                    parents_stack.append(node)

            # elif paragraph.style.name == "Normal":
            # elif style.startswith("Body Text"):
            else:  # Body Text
                if summarizer:
                    # accumulate whole section
                    section_text.append(paragraph.text)
                else:
                    selected_text = [run.text for run in paragraph.runs if text_filter(run)]
                    if selected_text:
                        if doc_file:
                            compiled_text = ' […] '.join(selected_text)
                            self._output_document.add_paragraph(compiled_text)
                        if freeplane_file:  # Freeplane
                            # TODO: use topic models
                            node_parent = parents_stack[-1].add_child(selected_text[0])
                            node_parent._node.attrib["STYLE"] = 'fork'
                            for segment in selected_text[1:]:
                                node = node_parent.add_child(segment)
                                node._node.attrib["STYLE"] = 'fork'

        # summarize residual last section if needed
        if summarizer and section_text:
            compiled_text = '\n'.join(section_text)
            append_summary(compiled_text)
            section_text = []

        # for h, t in zip(headings, texts):
        #     print(h, t)
        # print('\n'.join(styles))

        if doc_file:
            self._output_document.save(doc_file)
        if freeplane_file:
            self._output_midmapping.save(freeplane_file, encoding='utf-8')
        # if not doc_file and not freeplane_file:


def process_file(input_file_path: Path, actions: dict,
                 doc_path: Path = None, freeplane_path: Path = None, file_tag='outline'):
    assert input_file_path.exists(), 'input file not found'
    outliner = Outliner(input_file_path)

    if doc_path and doc_path.is_dir():
        doc_path = Path(doc_path, f'{input_file_path.stem}_{file_tag}.docx')

    if freeplane_path and freeplane_path.is_dir():
        freeplane_path = Path(freeplane_path,  f'{input_file_path.stem}_{file_tag}.mm')

    print(input_file_path, ' --> ', doc_path, freeplane_path)

    outliner.extract_marks(doc_file=doc_path, freeplane_file=freeplane_path, **actions)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="""Reads a `docx` word document and outputs a new document that preserves document structure (headings) but replaces content with: 
+ An automatic summary (first detects language, then summarizes)
A compilation of bold/highlighted segments (they were annotated by the user first maybe because they represent important terms or ideas
    """)
    # positional
    parser.add_argument('input', metavar='input file', type=str,
                        help='input path (file or folder) for highlighting compilation or automatic summarization')
    # parser.add_argument('output', metavar='output file', type=str, help='output path (file or folder)')

    # OUTPUT
    parser.add_argument('-d', '--docx', type=str, help='Specify word document output path')
    parser.add_argument('-f', '--freeplane', type=str, help='Specify freeplane document output path')
    parser.add_argument('-r', '--recursive', action='store_true',
                        help='Use it along a folder input and output: it makes available all files within subfolders')

    # conversion options
    parser.add_argument('-b', '--bold', action='store_true', help='filter bold content')
    parser.add_argument('-u', '--underline', action='store_true', help='filter underlined content')
    parser.add_argument('-s', '--summary', type=str, help='creates a summary, it requires language code, eg.: en')
    # optional & mutually exclusive
    # parser.add_argument("-a", "--action", type=str, default='marks', choices=["marks", "summary"], help="Choose action to process document content")

    args = parser.parse_args()
    # print(args)

    assert args.docx or args.freeplane, 'A docx or freeplane output file/folder must be provided'
    input_path = Path(args.input)
    assert input_path.exists(), 'input file/folder not found'

    if args.summary:
        summarizer = load_pipeline(language=args.summary)
        actions = {'summarizer': summarizer}
        file_tag = 'summary'
    elif args.bold or args.underline:
        actions = {'select_bold': args.bold, 'select_underline': args.underline}
        file_tag = 'annotation'
    else:
        raise SyntaxError('some arguments must be provided to indicate action on content: (-b -u | -s <lang>)')

    if input_path.is_file():
        files = [input_path]
    elif input_path.is_dir():
        if args.docx:
            assert Path(args.docx).is_dir()
        if args.freeplane:
            assert Path(args.freeplane).is_dir()

        # get input files
        if args.recursive:
            files = input_path.rglob('*.docx')
        else:
            files = input_path.glob('*.docx')
        assert files, 'Not .docx files found in ' + str(path2.resolve())

    for file_path in files:
        process_file(file_path, actions,
                     Path(args.docx) if args.docx else None, Path(args.freeplane) if args.freeplane else None,
                     file_tag)
