==================================
rstdoc-xgrid-project Version 1.8.4 
==================================

---------------------
Xgrid's Version 1.0.0
---------------------

This is the modified version of the original open-source `rstdoc <https://github.com/rstdoc/rstdoc>`_ python library, modified to suit the company's internal documentation needs.

Notes 
*****

* See `Original background and documentation <https://rstdoc.readthedocs.io/en/latest/>`__.

* This python package supports working with `RST <http://docutils.sourceforge.net/docs/ref/rst/restructuredtext.html>`_ as documentation format without depending on Sphinx.


* ``pip install rstdoc`` installs:

+-----------+-------------------+--------------------------------------------+
| Module    | CLI Script        | Description                                |
+===========+===================+============================================+
| dcx       | `rstdcx`_, rstdoc | create ``.tags``, labels and links         |
+-----------+-------------------+--------------------------------------------+
| fromdocx  | `rstfromdocx`_    | Convert DOCX to RST using Pandoc           |
+-----------+-------------------+--------------------------------------------+
| listtable | `rstlisttable`_   | Convert RST grid tables to list-tables     |
+-----------+-------------------+--------------------------------------------+
| untable   | `rstuntable`_     | Converts certain list-tables to paragraphs |
+-----------+-------------------+--------------------------------------------+
| reflow    | `rstreflow`_      | Reflow paragraphs and tables               |
+-----------+-------------------+--------------------------------------------+
| reimg     | `rstreimg`_       | Rename images referenced in the RST file   |
+-----------+-------------------+--------------------------------------------+
| retable   | `rstretable`_     | Transforms list tables to grid tables      |
+-----------+-------------------+--------------------------------------------+

