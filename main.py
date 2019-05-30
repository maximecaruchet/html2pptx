#!/usr/bin/env python3

import re
import requests
from bs4 import BeautifulSoup


def html_to_pptx(url, css_selector):
    r = requests.get(url)
    url_string = r.text
    slides = html_to_slides(url_string, css_selector)
    slides_to_pptx(slides)


def html_to_slides(html_string, css_selector):
    soup = BeautifulSoup(html_string, 'html.parser')
    useful_content = soup.select(css_selector)
    slides = []
    for parent_content_tag in useful_content[0].children:
        if parent_content_tag.name is not None:
            slide_content = html_to_slide(parent_content_tag)
            slides.append(slide_content)
    return slides


def html_to_slide(parent_tag):
    return parse_tag_contents(parent_tag)


def parse_tag_contents(tag):
    tag_data = []
    for children_content_tag in tag.children:
        # Go through all children tags
        if children_content_tag.name is not None:
            # Just handle valid tags
            if children_content_tag.name == "img":
                # If we have an image, get the "src" link
                tag_data.append("img_src:" + children_content_tag["src"])
            elif children_content_tag.string is not None:
                # If we have only one string, return it
                if children_content_tag.string.strip() != "":
                    tag_data.append(children_content_tag.string.strip())
            else:
                # Get direct text elements from tag even if there are children elements with text inside
                # (but do not get the text from the children)
                direct_tag_strings = children_content_tag.find_all(string=True, recursive=False)
                sanitized_direct_tag_strings = []
                for string in direct_tag_strings:
                    sanitized_string = string.strip()
                    if re.match("^\\[if mso \\| IE\\].*", sanitized_string):
                        sanitized_string = ""
                    if sanitized_string != "":
                        sanitized_direct_tag_strings.append(sanitized_string)

                # Get direct text elements from tag even if there are children elements with text inside
                # (this time, we get the text from the children)
                recursive_tag_strings = children_content_tag.find_all(string=True, recursive=True)
                sanitized_recursive_tag_strings = []
                for string in recursive_tag_strings:
                    sanitized_string = string.strip()
                    if re.match("^\\[if mso \\| IE\\].*", sanitized_string):
                        sanitized_string = ""
                    if sanitized_string != "":
                        sanitized_recursive_tag_strings.append(sanitized_string)

                # If we have some direct text elements, then we are in a case of some formatted text nested within other
                # text tags, then just extract the whole text (direct and from children), return it,
                # and stop the recursion by going directly to the next element
                if len(sanitized_direct_tag_strings) > 0:
                    tag_data.append(" ".join(sanitized_recursive_tag_strings))
                    continue

                # If we are not in the case of nested text, just do a recursive call
                # to get the contents of the children tag
                tag_data.extend(parse_tag_contents(children_content_tag))
    return tag_data


def slides_to_pptx(slides):
    for slide in slides:
        print("============================================ NEW SLIDE ============================================")
        for slide_data in slide:
            img_found = re.search("^img_src:(.*)$", slide_data)
            if img_found:
                print("IMAGE LINK:", img_found.group(1))
            else:
                print(slide_data)
        print("============================================ END SLIDE ============================================")


html_to_pptx("", "")
