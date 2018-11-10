"""
Functions intended to be called from spreadsheet

Note arguments extracted from the spreadsheet are always strings, so functions
need to convert into other data types as needed.
"""
import subprocess
import requests
from selenium import webdriver

# Dictionary intended to be used to share data between functions
share = {}

def add(a, b):
    # Args from spreadsheet are str, so need to convert to numbers
    return float(a) + float(b)

def fibonacci(iters=10):
    """Get Fibonacci series and return as csv string"""
    iters = int(iters)
    a, b = 1, 2
    results = [a, b]
    for i in range(iters):
        a, b = b, a+b
        results.append(b)
    #Convert to csv string for return to spreadsheet cell
    results = [str(r) for r in results]
    results = ", ".join(results)
    return results

def ping(url):
    """Ping website once using subprocess module and ping"""
    # -c flag might not be valid on Windows
    return subprocess.check_output(["ping", "-c", "1", url])

def url_status_code(url):
    """Get website status code using request"""
    r = requests.get(url)
    return r.status_code

def selenium_get_website_headings(url, tag='h4'):
    """Use webdrive to open webpage and extract heading text"""
    driver = webdriver.Firefox()
    driver.get(url)
    headlines = [e.text for e in driver.find_elements_by_tag_name(tag)]
    driver.close()
    return ", ".join(headlines)
    
# Example of two functions that share data
def read_file(filename, share_key="data_file"):
    with open(filename, "r") as infile:
        share[share_key] = infile.read()
    
def count_occurs(item="*", share_key="data_file"):
    return share.get(share_key,"").count(item)
