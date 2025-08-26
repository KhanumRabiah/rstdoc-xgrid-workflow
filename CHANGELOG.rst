=========
CHANGELOG
=========

202508 - v1.8.3
===============

Changes made are:

- In the ``fromdoc.py``, ``dcx.py``, files ``.rest`` is changed to ``.rst``
- In the ``reimg.py`` file ``-g`` flag is updated to fix the backslash issue in the image definitions
- In the ``dcx.py`` copyright and conf.py file updated for correct reflection of project effort
- Removed ``dcx.py`` file write feature into output folder function in ``fromdoc.py`` file.
- Modified ``fromdoc.py`` to include admoninitons directives.
- Modified ``fromdoc.py`` to include correct detection of code-block to convert it to respective directives.
- Multiple ``fromdoc.py`` to include the logic for the conversion of multiple ``.docx`` files to multiple ``.rst`` files.

.. TODO
.. ====

.. - use docutils make_id() to create the external target in the _links_xxx.rst files

.. - test tags on vscode, and add to docs


20230122 - v1.8.2
=================

- allow atx headers starting with a number of # chars
- run and fix tests with new pandoc, docutils, sphinx version

20201231- v1.8.1
================

- ``--version`` option
- fix regression tests
- more flexible ``pdtid()`` and ``pdtAAA()``

20191124 - v1.8.0
=================

No changelog yet.
Some later entries from git log:

- fix tests
- ``--rstrest`` to have sample project with .rst main and .rest
- use txdir
- ``/_links_sphinx.rst`` to search up dir
- ``--ipdt`` and ``--over`` sample project added
- support links accross directories
- allow control over file name in temporary directory
- targets with relative path
- do all stpl files in the tree, not just those in doc dir
- gen file can now also be python code
- cairosvg is brittle on windows. use inkscape in case of problems
- fix test, as pandoc 2.7 produces simple table now, instead of grid table
- added pytest.ini to suppress deprecation warnings
- grep and keyword search
- generated files for readthedocs
- work without sphinx_bootstrap_theme

20190228 - v1.6.8
=================

no changelog yet
