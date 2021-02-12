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
			> convert pdf to images
			> convert images to DICOM study
			or 
			> convert pdf to dicom pdf file using series information from the DICOM SR file
		> possibly send file using dcmsend
		

## The config file
These are the config options with default values and description (sorted by importance descending):

    "template_path": "template.docx", 				// the path of the word or text template (any path without docx suffix will be treated as text file) 
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
    "target": "dcm_images",                         // OPTIONAL, the output that should be generated: one of "xml" (generate only xml), "template", "pdf", "dcm_pdf", "dcm_images"
    "output_dir": null,                             // OPTIONAL, a directory where temporary files are written, if set to null a temp dir will be used.set this to a known path to keep output files, e.g. when you need to import files from that location
	"output_file_name": null						// OPTIONAL, the file name for output files (extension will be appended)
    "temp_dir": null,                               // OPTIONAL, a directory where temporary files are written, if set to null a temp dir will be used. set this to a known path to keep intermediate files, e.g. for debugging   
    "additional_paths": [] 							// OPTIONAL, additional system paths to be added upon running (should point to poppler and dcmtk directory)
    "oid_root": null,                               // OPTIONAL, The organization root for generating DICOM UIDs, see: http://dicom.nema.org/dicom/2013/output/chtml/part05/chapter_9.html, if none some random will be used
    "dsr2xml_exe_additional_options": [
        "-Ee",
        "-Ec"
    ], 						                        // OPTIONAL, additional options for the xml conversion, see https://support.dcmtk.org/docs/dsr2xml.html
    "pdf2dcm_exe_additional_options": [], 	        // OPTIONAL, additional options for the pdf2dcm conversion, see https://support.dcmtk.org/docs/pdf2dcm.html
    "img2dcm_exe_additional_options": [
        "--no-checks" ], 						    // OPTIONAL, additional options for the img2dcm conversion, see https://support.dcmtk.org/docs/img2dcm.html
    "dcmsend_exe_additional_options": [],           // OPTIONAL, additional options for the dcmsend, see https://support.dcmtk.org/docs/dcmsend.html
    "dcm_send_ip": null, 							// OPTIONAL, dcmsend ip destination, HINT: if this is null, no dcmsend command will be issued
    "dcm_send_port": null,  						// REQUIRED ONLY IF dcm_send_ip != null, dcmsend port
    "dcm_send_dcm_sr": false,                       // OPTIONAL, whether to send the original SR report also,

