# Usage
The application is used for creating (Word) flyers and html texts for Motiv8s calender.

# Configurable data
## url_root
The current url root is: https://motiv8s.org/motiv8s. This can be changed in the (local) file resources\web.csv.

## Caller pictures
Caller pictures (jpg format) are stored at <url_root>/caller_images/bordered/.

A list of caller pictures is created by the script <url_root>/scripts/get_callers.php.

New pictures can be uploaded by FTP.

## Dance schemas
Dance schemas are json data files, stored at <url_root>/dance_schemas.

New schemas can be uploaded by FTP.

## Texts
In the (local) file resources\texts_se.csv and resources\texts_en.csv are swedish and english texts stored in key;value pairs
Values can be edited, but keys cannot be changed without changing the program code. 
