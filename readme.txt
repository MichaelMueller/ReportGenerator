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
    "template_path": "template.docx", 				// the path of the word or text template (any path without docx suffix will be treated as text file)
    "dsr2xml_exe": "dsr2xml", 						// the path of the dcmtk dsr2xml tool (leave it if you have the dcmtk binary directory in your PATH)
    "pdf2dcm_exe": "pdf2dcm", 						// the path of the dcmtk pdf2dcm tool (leave it if you have the dcmtk binary directory in your PATH)
    "dcmsend_exe": "dcmsend", 						// the path of the dcmtk dcmsend tool (leave it if you have the dcmtk binary directory in your PATH)
    "dcm_send_ip": null, 							// dcmsend ip destination, HINT: if this is null, no dcmsend command will be issued
    "dcm_send_port": null,  						// dcmsend port
    "keep_temp_files": false, 						// whether to keep temp files (useful for debugging)
    "output_dicom_xml_file": null,					// where the intermediate xml file will be placed, if this is null a temp file will be used
    "quit_after_xml_file_creation": false,			// set this to true if you need to just create the xml file
    "output_template_file": null,					// where the filled (!) template file will be written. if this is null a temp file will be used.
    "output_dicom_pdf_file": "report09.pdf.dcm", 	// the output dcm file, if this is null a temp name will be used
    "skip_pdf_file_creation": false, 				// if True no pdf will be created. use this if you just want to convert to text files
    "rules": [ 										// there can be multiple rules: a rule is a set of instructions for text extraction and replacement
        {
            "name": "$findings$", 					// denotes to the exact placeholder name inside the template document (text or word file)
            "concat_string": "\n", 					// if multiple texts are extracted by the xpath_expressions, this will be the glue to make it one string
            "xpath_expressions": [ 					// there can be multiple xpath expressions to extract text parts from the dicom sr xml
                "/report/document/content/container/text[concept/meaning[contains(text(), \"Finding\")]]/value/text()"
            ],
            "replacements": { 						// optional replacement strings, will be applied to all texts (!) extracted by the xpath_expressions
                "<BR>": "\n"
            }
        }
    ],
    "additional_paths": [] 							// additional system paths to be added upon running
}