## Pollinate
POLLINATE is a [Python](https://python.org) pingback via an HTTP GET request.  There are multiple ways to identify the source IP address for an opened Microsoft Office document ranging from a macro to embedded single pixels images. This method takes advantage of the XML-based file format for [Word Documents](https://docs.microsoft.com/en-us/deployoffice/compat/office-file-format-reference) and is based off concepts from [Rhino Security Labs](https://github.com/RhinoSecurityLabs/Security-Research/blob/master/tools/ms-office/subdoc-injector/subdoc_injector.py) and [tifkin-](https://gist.github.com/tifkin-/a29fb9b88f029216d192).


## Features
* Cross-platform (.docx [Word Documents](https://docs.microsoft.com/en-us/deployoffice/compat/office-file-format-reference))

## Dependencies
* [Python](https://python.org) 3.6+

## Usage
Add a pingback via HTTP to a Microsoft Office document

Examples
```
python3 pollinate.py -f somedocument.docx -u http://127.0.0.1:8080 -o test.docx

python3 pollinate.py -f somedocument.docx -u http://example.com:8080 -o test.docx

python3 pollinate.py -f somedocument.docx  -u 'http://<FQDN>:<PORT>/index.html?name=value' -o test.docx

```

Capturing the opened document's public IP address, one can leverage a webserver or simple python server for prototyping.

Python 2
```
python -m SimpleHTTPServer 8080
```
Python 3
```
python3 -m http.server 8080
```

## OPSEC
* Constants for signature-based detection based off:
    * The ID variables (Id="rId1337") in both ```relsData``` and ```drawing```
    * Pingback URL and or parameters
* This POC only uses HTTP
