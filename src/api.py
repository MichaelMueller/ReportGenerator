import abc
import hashlib
import logging
import os
import platform
import shutil
import subprocess
import sys
from typing import List, Optional

def setup_logging(log_level=logging.INFO, log_file=None):
    class InfoFilter(logging.Filter):
        def filter(self, rec):
            return rec.levelno in (logging.DEBUG, logging.INFO, logging.WARNING)

    h1 = logging.StreamHandler(sys.stdout)
    h1.setLevel(logging.DEBUG)
    h1.addFilter(InfoFilter())
    h2 = logging.StreamHandler(sys.stderr)
    h2.setLevel(logging.ERROR)

    handlers = [h1, h2]
    kwargs = {"format": "%(asctime)s,%(msecs)d %(levelname)-8s [%(filename)s:%(lineno)d] %(message)s",
              "datefmt": '%Y-%m-%d:%H:%M:%S', "level": log_level}

    if log_file:
        h1 = logging.FileHandler(filename=log_file)
        h1.setLevel(logging.DEBUG)
        handlers = [h1]

    kwargs["handlers"] = handlers
    logging.basicConfig(**kwargs)


def sha256_sum(file_path):
    h = hashlib.sha256()
    b = bytearray(128 * 1024)
    mv = memoryview(b)
    with open(file_path, 'rb', buffering=0) as f:
        for n in iter(lambda: f.readinto(mv), 0):
            h.update(mv[:n])
    return h.hexdigest()


def get_int(str):
    try:
        value = int(str)
        return value
    except ValueError:
        return str


class FileProcessor:
    @abc.abstractmethod
    def process(self, full_path, basename, file_name, ext, dir_name):
        pass


class TextProcessor:
    @abc.abstractmethod
    def process(self, text, full_path, basename, file_name, ext, dir_name):
        pass


class FileIndex:
    @abc.abstractmethod
    def is_indexed(self, full_path):
        return False


class DirParser:
    def __init__(self, directory, file_processor, extensions=[]):
        self.directory = directory  # type: Optional[str]
        self.file_processor = file_processor  # type: FileProcessor
        self.extensions = extensions  # type: List[Optional[str]]

    def run(self):
        for root, dirs, files in os.walk(self.directory):
            for basename in files:
                file_name, ext = os.path.splitext(basename)
                ext = ext[1:].lower()
                if len(self.extensions) > 0 and ext not in self.extensions:
                    continue
                path = os.path.join(root, basename)

                self.file_processor.process(path, basename, file_name, ext, root)


class TextractProcessor(FileProcessor):
    def __init__(self, text_processor, index):
        self.text_processor = text_processor  # type: TextProcessor
        self.index = index  # type: FileIndex

    def process(self, full_path, basename, file_name, ext, dir_name):
        logger = logging.getLogger(__name__)
        try:
            if self.index.is_indexed(full_path):
                logger.info("skipping already indexed file {}".format(full_path))
                return

            text = textract.process(full_path)
            self.text_processor.process(text.decode('unicode_escape'), full_path, basename, file_name, ext, dir_name)
        except Exception as e:
            logger.warning("error processing file {}: {}".format(full_path, str(e)))


class TesseractProcessor(FileProcessor):
    def __init__(self, text_processor, index):
        self.text_processor = text_processor  # type: TextProcessor
        self.index = index  # type: FileIndex

    def process(self, full_path, basename, file_name, ext, dir_name):
        logger = logging.getLogger(__name__)
        try:
            if self.index.is_indexed(full_path):
                logger.info("skipping already indexed file {}".format(full_path))
                return

            if ext == "pdf":
                images = convert_from_path(full_path)
            else:
                images = [Image.open(full_path)]

            text = ""
            for image in images:
                text = text + pytesseract.image_to_string(image)
        except Exception as e:
            logger.warning("error processing file {}: {}".format(full_path, str(e)))


class WhooshIndexer(TextProcessor):
    def __init__(self, index_path, rebuild, commit_after_num_files=50):
        self.index_path = index_path
        self.rebuild = rebuild
        self.commit_after_num_files = commit_after_num_files
        # state
        self.ix_writer = None
        self.index_cleared = False
        self.file_num = 0

    def process(self, text, full_path, basename, file_name, ext, dir_name):

        logger = logging.getLogger(__name__)
        # sha256 = sha256_sum(full_path)
        self.assert_index_writer()
        logger.info("indexing file {}".format(full_path))
        self.ix_writer.add_document(title=basename, path=full_path, file_id=full_path, content=text, textdata=text,
                                    dir_name=dir_name)
        self.file_num = self.file_num + 1
        if self.commit_after_num_files is not None and self.file_num % self.commit_after_num_files == 0:
            self.ix_writer.commit()
            self.ix_writer = None

    def assert_index_writer(self):

        logger = logging.getLogger(__name__)
        if self.ix_writer is not None:
            return

        if self.rebuild and os.path.exists(self.index_path) and self.index_cleared == False and exists_in(
                self.index_path):
            logger.info("deleting index at {}".format(self.index_path))
            shutil.rmtree(self.index_path)
            self.index_cleared = True

        if os.path.exists(self.index_path):
            self.ix_writer = open_dir(self.index_path).writer()
        else:
            os.makedirs(self.index_path, 0o777, True)
            schema = Schema(title=TEXT(stored=True), path=TEXT(stored=True), dir_name=TEXT(stored=True),
                            file_id=ID(stored=True), content=TEXT, textdata=TEXT(stored=True))
            ix = create_in(self.index_path, schema)
            self.ix_writer = ix.writer()

    def __del__(self):
        if self.ix_writer:
            self.ix_writer.commit()


class WooshIndex(FileIndex):
    def __init__(self, index_path):
        self.index_path = index_path
        # state
        self.indexed_paths = None

    def is_indexed(self, full_path):
        self.assert_indexed_paths()
        return full_path in self.indexed_paths

    def assert_indexed_paths(self):
        if self.indexed_paths is not None:
            return

        self.indexed_paths = set()

        if exists_in(self.index_path) == False:
            return

        ix = open_dir(self.index_path)
        self.indexed_paths = set()
        with ix.searcher() as searcher:
            for fields in searcher.all_stored_fields():
                indexed_path = fields['path']
                self.indexed_paths.add(indexed_path)


class DirOcr:

    def index(self, directory, index_path, rebuild):
        whoosh_indexer = WhooshIndexer(index_path, rebuild)
        whoosh_index = WooshIndex(index_path)

        file_processor = TextractProcessor(whoosh_indexer, whoosh_index)
        extensions = []

        dir_parser = DirParser(directory, file_processor, extensions)

        dir_parser.run()

    def interactive_search(self, index_path, num_docs):
        while True:
            query_str = input("query_string: ")
            if query_str:
                self.search(index_path, query_str, num_docs)

    def open_file(self, full_path):
        if platform.system() == 'Darwin':  # macOS
            subprocess.call(('open', full_path))
        elif platform.system() == 'Windows':  # Windows
            os.startfile(full_path)
        else:  # linux variants
            subprocess.call(('xdg-open', full_path))

    def search(self, index_path, query_str, num_docs):
        ix = open_dir(index_path)

        with ix.searcher(weighting=scoring.Frequency) as searcher:
            query = QueryParser("content", ix.schema).parse(query_str)
            results = searcher.search(query, limit=num_docs)

            if len(results) == 0:
                print("!!!No Results!!!")
            else:
                for i, result in enumerate(results):
                    print("{}: {} in {}".format(i + 1, result['title'], result['dir_name']))

                action = get_int(input("next search (ENTER), open file (<number>), open all files (a), quit(q): "))
                if action:
                    if action == "a":
                        for result in results:
                            self.open_file(result['path'])
                    elif isinstance(action, int) and 0 < action <= len(results):
                        self.open_file(results[action-1]['path'])
                    elif action == "q":
                        return

    def __call__(self):
        app = QApplication(sys.argv)
        ex = App()
        sys.exit(app.exec_())
