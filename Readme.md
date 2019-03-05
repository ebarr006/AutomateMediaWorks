# Automated Media Works Request Tool
Written by Emilio Barreiro

### Overview
Part of planning for on-campus events, the UCR Undergraduate Admissions Fiscal Team needs to submit numerous work order type requests to prepare media in classrooms. Often times, there is a long list of (50+) rooms with different configurations, request times, and media needs. The Media Works site only allows users to request media for a single room per form submission, making the reservation process very repetitive and time consuming.

I wrote a tool to automate the submission process using Selenium Web Driver. The tools first loads the event data, equipment list, and required rooms from an Excel file (.xlsx). Next it drives a web browser to UCR's Media Works site. Then it will fill out a submission form for each room location and corresponding equipment list. 
