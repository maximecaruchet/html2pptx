# html2pptx

# How to use

This program requires Python 3 in order to work  
Install all requirements from `requirements.txt` by doing `pip install -r requirements.txt`

Once everything is installed you can run html2pptx by executing the `main.py` file

You should be able to access the `index.html` page at the address where the server is located

The page expects two arguments, one URL and one CSS selector  
When you click on the submit button, html2pptx will get the URL contents and extract the HTML
elements targeted by the CSS selector.  
The HTML element targeted must be a parent element with multiple children elements as siblings:

```html
<div> <!-- <== This is the element you should target -->
    <p>...</p>
    <p>...</p>
    <p>...</p>
    <p>...</p>
    <p>...</p>
</div>
```

Every direct child element of the parent element will become a slide and its contents will be
all the children of the direct child element  
As for the contents and the layout, the service will try its best to get all images and text from the page
and rearrange them with a slide layout depending on various elements (number of images, text length, ...)
