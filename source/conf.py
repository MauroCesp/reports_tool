# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

import os
import sys

#Provide the path the the code is
# I separated the code from the source, I need to provide the code (back directory)
sys.path.insert(1, os.path.abspath('../code'))

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

project = 'Reports tool'
copyright = '2023, Mauro Céspedes Araya'
author = 'Mauro Céspedes Araya'
release = '1.1'

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

# This extension we need to add it to generate 
extensions = [
   'sphinx.ext.duration',
   'sphinx.ext.doctest',
   'sphinx.ext.autodoc',
   'sphinx.ext.autosummary',
]

templates_path = ['_templates']
exclude_patterns = []



# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output
# nature, calssic, sphinxdocs, bizstyle
html_theme = 'nature'
html_static_path = ['_static']

html_sidebars = { '**': ['globaltoc.html', 'relations.html',
        'sourcelink.html', 'searchbox.html'], }