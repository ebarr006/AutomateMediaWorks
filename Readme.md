# Automated Media Works Request Tool
Written by Emilio Barreiro

### Overview
Part of planning for on-campus events, the UCR Undergraduate Admissions Fiscal Team needs to submit numerous work order type requests to prepare media in classrooms. Often times, there is a long list of (50+) rooms with different configurations, request times, and media needs. The Media Works site only allows users to request media for a single room per form submission, making the reservation process very repetitive and time consuming.

I wrote a tool to automate the submission process using Selenium Web Driver. The tool loads the room configurations from a json file and, with one browser, creates new requests to fill out and submit. The next step for this project will be having the tool read from an excel file (.xlsx).
