import json
import logging
import os
import subprocess
import sys
import tempfile
import importlib
import importlib_metadata
from win32com import client

from contextlib import contextmanager
from typing import List, Optional, Dict

# import docx2pdf
import lxml.etree as ET
from docx import Document


class DataObject:

    def from_dict(self, data: Dict):
        for key, value in data.items():
            setattr(self, key, value)

    def to_dict(self):
        return self.__dict__

    def validate(self):
        return None


class Rule(DataObject):

    @staticmethod
    def create_from_dict(data: Dict):
        rule = Rule()
        rule.from_dict(data)
        return rule

    def __init__(self, name=None, concat_string="\n", xpath_expressions=[], replacements={}):
        self.name = name
        self.concat_string = concat_string
        self.xpath_expressions = xpath_expressions
        self.replacements = replacements

    def validate(self):
        error = None
        if not self.name or not self.concat_string or not self.xpath_expressions:
            error = "this values may not be empty: name, concat_string, xpath_expressions"
        return error


class Config(DataObject):
    @staticmethod
    def create_from_dict(data: Dict):
        config = Config()
        config.from_dict(data)
        return config

    def __init__(self, template_path=None, dsr2xml_exe="dsr2xml", rules=[]):
        self.template_path = template_path  # type: Optional[str]
        self.dsr2xml_exe = dsr2xml_exe  # type: Optional[str]
        self.pdf2dcm_exe = "pdf2dcm"
        self.dcm_send_ip = None
        self.dcm_send_port = None
        self.keep_temp_files = False
        self.output_dicom_pdf_file = None
        self.rules = rules  # type: List[Rule]

    def validate(self):
        error = ""
        if not self.template_path or not self.dsr2xml_exe or not self.rules or not self.pdf2dcm_exe:
            error = "this values may not be empty: template_path, dsr2xml_exe, rules, pdf2dcm_exe"
        for idx, rule in enumerate(self.rules):
            rule_error = rule.validate()
            if rule_error:
                if error:
                    error = error + "\n"
                error = error + "error in rule " + str(idx + 1) + ": " + rule_error

        return "config error: " + error if error else None

    def from_dict(self, data: Dict):
        if "rules" in data.keys():
            rules = []
            for rule_data in data["rules"]:
                rules.append(Rule.create_from_dict(rule_data))
            data["rules"] = rules
        super().from_dict(data)

    def to_dict(self):
        data = super().to_dict()
        for idx, rule in enumerate(data["rules"]):
            data["rules"][idx] = rule.to_dict()
        return data


def doc2pdf(doc_name, pdf_name):
    """
    :word to pdf
    :param doc_name word file name
    :param pdf_name to_pdf file name
    """
    word = client.DispatchEx("Word.Application")
    if os.path.exists(pdf_name):
        os.remove(pdf_name)
    worddoc = word.Documents.Open(doc_name, ReadOnly=1)
    worddoc.SaveAs(pdf_name, FileFormat=17)
    worddoc.Close()
    #return pdf_name


def setup_logging(log_level=logging.INFO, log_file=None):
    class InfoFilter(logging.Filter):
        def filter(self, rec):
            return rec.levelno in (logging.DEBUG, logging.INFO, logging.WARNING)

    h1 = logging.StreamHandler(sys.stdout)
    h1.flush = sys.stdout.flush
    h1.setLevel(logging.DEBUG)
    h1.addFilter(InfoFilter())
    h2 = logging.StreamHandler(sys.stderr)
    h2.flush = sys.stderr.flush
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


def quit(error):
    logger = logging.getLogger(__name__)
    logger.error(error)
    sys.exit(-1)


def run_cmd(*args, print_stdout=True):
    logger = logging.getLogger(__name__)
    cmd = ' '.join(args)
    logger.debug("running the following command: {}".format(cmd))
    result = subprocess.run(args, stdout=subprocess.PIPE)
    if result.returncode != 0:
        quit(
            "cmd \"{}\" failed with code {} the following output: {}. aborting.".format(cmd, str(result.returncode),
                                                                                        result.stderr))
    elif print_stdout:
        result = result.stdout.decode("utf-8")
        if result:
            logger.info(result)


def replace_in_docx(docx_path, data, output_docx_path):
    doc = Document(docx_path)
    for i in data:
        for p in doc.paragraphs:
            if p.text.find(i) >= 0:
                p.text = p.text.replace(i, data[i])
    # save changed document
    doc.save(output_docx_path)


def dump_config(dump_file, log_level, log_file):
    # logging
    setup_logging(log_level, log_file)
    logger = logging.getLogger(__name__)
    logger.info("dumping default config into {}".format(dump_file))

    # create default config
    rule = Rule("$findings$")
    rule.xpath_expressions.append(
        '/report/document/content/container/text[concept/meaning[contains(text(), "Finding")]]/value/text()')
    rule.replacements["<BR>"] = "\n"
    config = Config(template_path="../sample_data/template.docx")
    config.rules.append(rule)

    data = config.to_dict()
    with open(dump_file, 'w') as out_file:
        json.dump(data, out_file, indent=4)


@contextmanager
def suppress_stdout():
    with open(os.devnull, "w") as devnull:
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sys.stdout = devnull
        sys.stderr = devnull
        try:
            yield
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr


def generate_report(dcm_sr_path, config_file, log_level, log_file):
    # logging
    setup_logging(log_level, log_file)
    logger = logging.getLogger(__name__)

    # files need to be deleted
    try:
        with open(config_file) as json_file:
            data = json.load(json_file)
            config = Config.create_from_dict(data)
            error_str = config.validate()
            if error_str:
                quit(error_str)

        # GENERATE XML FILE
        with tempfile.NamedTemporaryFile(suffix=".xml", delete=not config.keep_temp_files) as tmp_file:
            sr_xml_file = tmp_file.name
        logger.info("converting DICOM SR {} to XML file {}".format(dcm_sr_path, sr_xml_file))
        run_cmd(config.dsr2xml_exe, dcm_sr_path, sr_xml_file)

        # EXTRACT AND CONCAT CONTENTS USING XPATH
        logger.info("retrieving contents from XML file {}".format(sr_xml_file))
        root = ET.parse(sr_xml_file)
        template_data = {}
        for rule in config.rules:
            text = ""
            for xpath_expression in rule.xpath_expressions:
                xpath_result = root.xpath(xpath_expression)
                print(xpath_result)
                if isinstance(xpath_result, List):
                    xpath_result = rule.concat_string.join(xpath_result)

                if xpath_result:
                    if text:
                        text = text + rule.concat_string
                    text = text + xpath_result
                    for search, replace in rule.replacements.items():
                        text = text.replace(search, replace)
            template_data[rule.name] = text
        logger.debug("template_data: {}".format(str(template_data)))

        # LOAD TEMPLATE AND SET CONTENTS ON NAMED PLACEHOLDERS
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=not config.keep_temp_files) as tmp_file:
            docx_tmp_file = tmp_file.name
        logger.info("replacing contents from template docx file {} into {}".format(config.template_path, docx_tmp_file))
        replace_in_docx(config.template_path, template_data, docx_tmp_file)

        # CONVERT TO PDF
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=not config.keep_temp_files) as tmp_file:
            pdf_tmp_file = tmp_file.name
        logger.info("converting file {} into pdf file {}".format(docx_tmp_file, pdf_tmp_file))
        with suppress_stdout():
            #docx2pdf.convert(docx_tmp_file, pdf_tmp_file)
            doc2pdf(docx_tmp_file, pdf_tmp_file)

        # CONVERT TO DICOM PDF
        if config.output_dicom_pdf_file:
            dcm_pdf_tmp_file = config.output_dicom_pdf_file
        else:
            with tempfile.NamedTemporaryFile(suffix=".dcm", delete=not config.keep_temp_files) as tmp_file:
                dcm_pdf_tmp_file = tmp_file.name
        logger.info("converting file {} into DICOM pdf file {}".format(pdf_tmp_file, dcm_pdf_tmp_file))
        run_cmd("pdf2dcm", pdf_tmp_file, dcm_pdf_tmp_file, "--series-from", dcm_sr_path)

        # SEND TO DICOM NODE
        if config.dcm_send_ip:
            logger.info("sending file {} to dicom node".format(dcm_pdf_tmp_file))
            # run_cmd("dcmsend", "localhost", "2727", dcm_sr_path)
            run_cmd("dcmsend", config.dcm_send_ip, config.dcm_send_port, dcm_pdf_tmp_file, print_stdout=False)

    except Exception as error:
        logger.exception(error)
