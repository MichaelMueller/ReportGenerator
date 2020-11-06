import json
import logging
import os
import shutil
import subprocess
import sys
import tempfile
from zipfile import ZipFile

import pdfkit
from pdf2image import pdf2image
from win32com import client

from contextlib import contextmanager
from typing import List, Optional, Dict

# import docx2pdf
import lxml.etree as ET
from docx import Document


class DataObject:

    def __init__(self):
        pass

    def from_dict(self, data: Dict):
        for key, value in data.items():
            setattr(self, key, value)

    def to_dict(self):
        data = {}
        for key, value in self.__dict__.items():
            data[key] = value
        return data

    def validate(self):
        return None


class Rule(DataObject):

    @staticmethod
    def create_from_dict(data: Dict):
        rule = Rule()
        rule.from_dict(data)
        return rule

    def __init__(self):
        super().__init__()
        self.name = None
        self.concat_string = "\n"
        self.xpath_expressions = []
        self.replacements = {}

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

    def __init__(self):
        super().__init__()
        self.template_path = None  # type: Optional[str]
        self.dsr2xml_exe_additional_options = ["-Ee", "-Ec"]  # type: Optional[List[str]]
        self.dcm_send_ip = None
        self.dcm_send_port = None
        self.keep_temp_files = False
        self.output_dicom_xml_file = None
        self.quit_after_xml_file_creation = False
        self.output_template_file = None
        self.output_dicom_pdf_file = None
        self.output_dicom_images_dir = "."
        self.create_dicom_from_images = True
        self.img2dcm_exe_additional_options = ["--no-checks"]
        self.skip_pdf_file_creation = False
        self.rules = []  # type: List[Rule]
        self.additional_paths = ["dcmtk-3.6.5-win64-dynamic/bin", "poppler-20.11.0/bin"]  # type List[str]

    def add_paths(self):
        for additional_path in self.additional_paths:
            os.environ["PATH"] += os.pathsep + additional_path

    def validate(self):
        error = ""
        if not self.template_path or not self.rules:
            error = "this values may not be empty: template_path, rules"
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
    # return pdf_name


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


def run_cmd(*args, print_stdout=False, exit_on_error=True):
    logger = logging.getLogger(__name__)
    cmd = ' '.join(args)
    logger.debug("running the following command: {}".format(cmd))
    if print_stdout:
        stderr = sys.stderr
        stdout = sys.stdout
    else:
        stderr = subprocess.PIPE
        stdout = subprocess.PIPE

    result = subprocess.run(args, stdout=stdout, stderr=stderr)
    if result.returncode != 0 and exit_on_error:
        quit(
            "cmd \"{}\" failed with code {} the following output: {}. aborting.".format(cmd, str(result.returncode),
                                                                                        result.stderr))
    return result.stdout.decode("utf-8").strip() if result.stdout else None


def replace_in_docx(docx_path, data, output_docx_path):
    doc = Document(docx_path)
    for i in data:
        for p in doc.paragraphs:
            if p.text.find(i) >= 0:
                p.text = p.text.replace(i, data[i])
    # save changed document
    doc.save(output_docx_path)


def replace_in_text_file(in_file, data: Dict, out_file):
    # Read in the file
    with open(in_file, 'r') as file:
        file_data = file.read()

    # Replace the target string
    for placeholder, new_text in data.items():
        file_data = file_data.replace(placeholder, new_text)

    # Write the file out again
    with open(out_file, 'w') as file:
        file.write(file_data)


def create_default_config():
    # create default config
    rule = Rule()
    rule.name = "$findings$"
    rule.xpath_expressions.append(
        '/report/document/content/container/text[concept/meaning[contains(text(), "Finding")]]/value/text()')
    rule.replacements["<BR>"] = "\n"
    config = Config()
    config.template_path = "report09_template.docx"
    config.output_dicom_pdf_file = "report09.pdf.dcm"
    config.rules.append(rule)
    return config


def dump_config_to_file(dump_file, config):
    data = config.to_dict()
    with open(dump_file, 'w') as out_file:
        json.dump(data, out_file, indent=4)


def dump_config(dump_file, log_level, log_file):
    # logging
    setup_logging(log_level, log_file)
    logger = logging.getLogger(__name__)
    logger.info("dumping default config into {}".format(dump_file))

    config = create_default_config()
    dump_config_to_file(dump_file, config)


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


def zipdir(path, ziph):
    # ziph is zipfile handle
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file))


def create_installer(log_level=logging.INFO, log_file=None):
    # logging
    setup_logging(log_level, log_file)
    logger = logging.getLogger(__name__)
    logger.info("creating installer package".format())

    logger.info("running git commands".format())
    run_cmd("git", "add", "-A", print_stdout=True, exit_on_error=False)
    run_cmd("git", "commit", "-m", "'installer commit'", print_stdout=True, exit_on_error=False)
    rev_hash = run_cmd('git', 'rev-parse', 'HEAD', print_stdout=False)
    logger.info("current git hash is {}".format(rev_hash))

    logger.info("going to src dir")
    dir_path = os.path.dirname(os.path.realpath(__file__))
    os.chdir(dir_path)
    app_name = "ReportGenerator"
    output_dir = "../build/output"
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)

    logger.info("creating pyinstaller")
    run_cmd("pyinstaller", "--name", app_name, "--noconfirm", "--onefile", "--console", "report_generator.py",
            "--log-level", "WARN",
            "--clean", "--workpath", "../build/tmp", "--distpath", output_dir, "--specpath", "../build/tmp")

    logger.info("copying additional files")
    base_dir = "../base"
    src_files = os.listdir(base_dir)
    for file_name in src_files:
        if file_name.startswith("offis") or file_name.startswith("image"):
            continue
        full_file_name = os.path.join(base_dir, file_name)
        if os.path.isfile(full_file_name):
            dest = os.path.join(output_dir, file_name)
            shutil.copy(full_file_name, dest)

    shutil.copyfile("../readme.txt", output_dir + "/readme.txt")
    shutil.copytree(base_dir+"/dcmtk-3.6.5-win64-dynamic", output_dir + "/dcmtk-3.6.5-win64-dynamic")
    shutil.copytree(base_dir+"/poppler-20.11.0", output_dir + "/poppler-20.11.0")

    os.chdir(output_dir)
    logger.info("creating test case files: report09")
    report09_config = create_default_config()
    dump_config_to_file("report09_config.json", report09_config)
    report09_batch = open(r'ReportGenerator_report09.bat', 'w+')
    report09_batch.write(app_name + '.exe report09.dcm report09_config.json\nCMD')
    report09_batch.close()

    logger.info("creating test case files: report10")
    report10_config = create_default_config()
    report10_config.output_dicom_pdf_file = "report10.pdf.dcm"
    report10_config.output_template_file = "report10.html"
    report10_config.template_path = "report10_template.html"
    report10_config.skip_pdf_file_creation = True
    report10_config.output_dicom_xml_file = "report10.xml"

    report10_config.rules = []
    report10_rule1 = Rule()
    report10_rule1.name = "$findings$"
    report10_rule1.concat_string = "<br>"
    report10_rule1.xpath_expressions.append(
        '/report/document/content/container/container/text[concept/meaning[contains(text(), "Finding")]]/value/text()')
    report10_config.rules.append(report10_rule1)

    report10_rule2 = Rule()
    report10_rule2.name = "$name$"
    report10_rule2.concat_string = " "
    report10_rule2.xpath_expressions.append("/report/patient/name/first/text()")
    report10_rule2.xpath_expressions.append("/report/patient/name/last/text()")
    report10_config.rules.append(report10_rule2)

    dump_config_to_file("report10_config.json", report10_config)
    report10_batch = open(r'ReportGenerator_report10.bat', 'w+')
    report10_batch.write(app_name + '.exe report10.dcm report10_config.json\nCMD')
    report10_batch.close()

    additional_file = open(r'current_git_hash.txt', 'w+')
    additional_file.write(rev_hash)
    additional_file.close()

    zip_file = ZipFile("../" + app_name + '.zip', 'w')
    zipdir(".", zip_file)
    # close the Zip File
    zip_file.close()

    shutil.rmtree("../tmp")


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
        config.add_paths()

        # GENERATE XML FILE
        if config.output_dicom_xml_file:
            sr_xml_file = config.output_dicom_xml_file
        else:
            with tempfile.NamedTemporaryFile(suffix=".xml", delete=not config.keep_temp_files) as tmp_file:
                sr_xml_file = tmp_file.name
        logger.info("converting DICOM SR {} to XML file {}".format(dcm_sr_path, sr_xml_file))
        run_cmd("dsr2xml", *config.dsr2xml_exe_additional_options, dcm_sr_path, sr_xml_file)
        if config.quit_after_xml_file_creation:
            logger.info("xml created. quit requested.")
            sys.exit(0)

        # EXTRACT AND CONCAT CONTENTS USING XPATH
        logger.info("retrieving contents from XML file {}".format(sr_xml_file))
        root = ET.parse(sr_xml_file)
        template_data = {}
        for rule in config.rules:
            text = ""
            for rule_idx, xpath_expression in enumerate(rule.xpath_expressions):
                xpath_result = root.xpath(xpath_expression)
                if isinstance(xpath_result, List):
                    xpath_result = rule.concat_string.join(xpath_result)

                if not isinstance(xpath_result, str):
                    quit("xpath did not produce text: \"{}\" in rule {}, index {}".format(xpath_expression, rule.name,
                                                                                          str(rule_idx)))
                elif len(xpath_result) == 0:
                    logger.warning(
                        "empty text for xpath \"{}\" in rule {}, index {}".format(xpath_expression, rule.name,
                                                                                  str(rule_idx)))

                else:
                    logger.info("result for xpath {}: {}".format(xpath_expression, xpath_result))
                    if text:
                        text = text + rule.concat_string
                    text = text + xpath_result
                    for search, replace in rule.replacements.items():
                        text = text.replace(search, replace)

            template_data[rule.name] = text
            logger.debug("template_data: {}".format(str(template_data)))

        # LOAD TEMPLATE AND SET CONTENTS ON NAMED PLACEHOLDERS
        _, file_extension = os.path.splitext(config.template_path)
        template_is_word = file_extension == ".docx"
        if config.output_template_file:
            filled_template_file = config.output_template_file
        else:
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=not config.keep_temp_files) as tmp_file:
                filled_template_file = tmp_file.name
        logger.info("replacing contents from template docx file {} into {}".format(config.template_path,
                                                                                   filled_template_file))
        if template_is_word:
            replace_in_docx(config.template_path, template_data, filled_template_file)
        else:
            replace_in_text_file(config.template_path, template_data, filled_template_file)

        # CONVERT TO PDF
        images = None
        pdf_tmp_file = None
        temp_dir = tempfile.TemporaryDirectory()
        if not config.skip_pdf_file_creation:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=not config.keep_temp_files) as tmp_file:
                pdf_tmp_file = tmp_file.name
            logger.info("converting file {} into pdf file {}".format(filled_template_file, pdf_tmp_file))
            with suppress_stdout():
                if template_is_word:
                    doc2pdf(filled_template_file, pdf_tmp_file)
                else:
                    pdfkit.from_file(filled_template_file, pdf_tmp_file)
            if config.create_dicom_from_images:
                images = pdf2image.convert_from_path(pdf_tmp_file, paths_only=True, output_folder=temp_dir.name,
                                                     fmt="jpg")

        # CONVERT TO DICOM
        dcm_files = []
        if images:
            for idx, image in enumerate(images):
                # Do something here
                if config.output_dicom_images_dir:
                    dcm_file = os.path.join(config.output_dicom_images_dir,
                                            os.path.splitext(dcm_sr_path)[0] + "_" + str(idx) + ".dcm")
                else:
                    with tempfile.NamedTemporaryFile(suffix=".dcm", delete=not config.keep_temp_files) as tmp_file:
                        dcm_file = tmp_file.name
                logger.info("converting image {} into DICOM file {}".format(image, dcm_file))
                run_cmd("img2dcm", "--series-from", dcm_sr_path, *config.img2dcm_exe_additional_options, image,
                        dcm_file, print_stdout=True)
                dcm_files.append(dcm_file)
        elif pdf_tmp_file:
            # CONVERT TO DICOM PDF
            if config.output_dicom_pdf_file:
                dcm_pdf_tmp_file = config.output_dicom_pdf_file
            else:
                with tempfile.NamedTemporaryFile(suffix=".dcm", delete=not config.keep_temp_files) as tmp_file:
                    dcm_pdf_tmp_file = tmp_file.name
            logger.info("converting file {} into DICOM pdf file {}".format(pdf_tmp_file, dcm_pdf_tmp_file))
            run_cmd("pdf2dcm", pdf_tmp_file, dcm_pdf_tmp_file, "--series-from", dcm_sr_path)
            dcm_files.append(dcm_pdf_tmp_file)

        # SEND TO DICOM NODE
        if len(dcm_files) > 0 and config.dcm_send_ip:
            for dcm_file in dcm_files:
                logger.info("sending file {} to dicom node".format(dcm_file))
                # run_cmd("dcmsend", "localhost", "2727", dcm_sr_path)
                run_cmd("dcmsend", config.dcm_send_ip, config.dcm_send_port, dcm_file,
                        print_stdout=False)

    except Exception as error:
        logger.exception(error)
