# ReportGenerator
A command line tool which takes a DICOM SR and converts it into a nicely formatted DICOM pdf.
dcmtk is used on the command line.

The basic workflow is
DICOM SR file 
	> XML file 
		> evaluate rules: (given in a config file)
			> extract texts using xpath expressions
			> optionally replace strings inside the texts	
				> concatenate texts mostly using line breaks (configurable)
					> open word file and replace placeholders with the concatenated text
		> save word file
		> convert word file to pdf
		> convert pdf to dicom pdf file using series information from the DICOM SR file
		> possibly send file using dcmsend
		

## The config file
{
    "template_path": "template.docx", // the path of the word template
    "dsr2xml_exe": "dsr2xml", // the path of the dcmtk dsr2xml tool
    "pdf2dcm_exe": "pdf2dcm", // the path of the dcmtk pdf2dcm tool
    "dcmsend_exe": "dcmsend", // the path of the dcmtk dcmsend tool
    "dcm_send_ip": null, // dcmsend ip destination, if this is null, no dcmsend command will be issued
    "dcm_send_port": null,  // dcmsend port
    "keep_temp_files": false, // whether to keep temp files (useful for debugging)
    "output_dicom_pdf_file": "report09.pdf.dcm", // the output dcm file, if this is none a temp name will be used
    "rules": [ // there can be multiple rules: a rule is a set of instructions for text extraction and replacement
        {
            "name": "$findings$", // denotes to the exact placeholder value inside the template document
            "concat_string": "\n", // if multiple texts are extracted by the xpath_expressions, this will be the glue to make it one string
            "xpath_expressions": [ // there can be multiple xpath expressions to extract text parts from the dicom sr xml
                "/report/document/content/container/text[concept/meaning[contains(text(), \"Finding\")]]/value/text()"
            ],
            "replacements": { // optional replacement strings, will be issued to all texts extracted by the xpath_expressions
                "<BR>": "\n"
            }
        }
    ]
}