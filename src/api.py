import abc
import hashlib
import logging
import os
import platform
import shutil
import subprocess
import sys
import tempfile
from typing import List, Optional
import subprocess

import docx2pdf
import lxml.etree as ET
from docx import Document


class Rule:
    def __init__(self, name, concat_string="\n"):
        self.name = name
        self.concat_string = concat_string
        self.xpath_expressions = []


class Config:
    def __init__(self, template_path, dsr2xml_exe="dsr2xml"):
        self.template_path = template_path  # type: Optional[str]
        self.dsr2xml_exe = dsr2xml_exe  # type: Optional[str]
        self.rules = []  # type: List[Rule]


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


def run_cmd(*args):
    logger = logging.getLogger(__name__)
    cmd = ' '.join(args)
    logger.debug("running the following command: {}".format(cmd))
    result = subprocess.run(args, stdout=subprocess.PIPE)
    if result.returncode != 0:
        logger.error(
            "cmd \"{}\" failed with code {} the following output: {}. aborting.".format(cmd, str(result.returncode),
                                                                                        result.stderr))
        sys.exit(-1)
    else:
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

def generate_report(dcm_sr_path, log_level, log_file):
    # logging
    setup_logging(log_level, log_file)
    logger = logging.getLogger(__name__)

    # files need to be deleted
    try:
        rule = Rule("$findings$")
        rule.xpath_expressions.append(
            '/report/document/content/container/text[concept/meaning[contains(text(), "Finding")]]/value/text()')
        config = Config(template_path="../sample_data/template.docx")
        config.rules.append(rule)

        # GENERATE XML FILE
        with tempfile.NamedTemporaryFile(suffix=".xml", delete=log_level != logging.DEBUG) as tmp_file:
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
            template_data[rule.name] = text
        logger.debug("template_data: {}".format(str(template_data)))

        # LOAD TEMPLATE AND SET CONTENTS ON NAMED PLACEHOLDERS
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=log_level != logging.DEBUG) as tmp_file:
            docx_tmp_file = tmp_file.name
        logger.info("replacing contents from template docx file {} into {}".format(config.template_path, docx_tmp_file))
        replace_in_docx(config.template_path, template_data, docx_tmp_file)

        # CONVERT TO PDF
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=log_level != logging.DEBUG) as tmp_file:
            pdf_tmp_file = tmp_file.name
        logger.info("converting file {} into pdf file {}".format(docx_tmp_file, pdf_tmp_file))
        docx2pdf.convert(docx_tmp_file, pdf_tmp_file)

        # CONVERT TO DICOM PDF
        with tempfile.NamedTemporaryFile(suffix=".dcm", delete=log_level != logging.DEBUG) as tmp_file:
            dcm_pdf_tmp_file = tmp_file.name
        logger.info("converting file {} into pdf file {}".format(pdf_tmp_file, dcm_pdf_tmp_file))
        run_cmd("pdf2dcm", pdf_tmp_file, dcm_pdf_tmp_file, "--series-from", dcm_sr_path)

        # SEND TO DICOM NODE
        logger.info("sending file {} to dicom node".format(dcm_pdf_tmp_file))
        run_cmd("dcmsend", "localhost", "2727", dcm_sr_path)
        run_cmd("dcmsend", "localhost", "2727", dcm_pdf_tmp_file)

    except Exception as error:
        logger.exception(error)

