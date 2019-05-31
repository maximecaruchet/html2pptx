#!/usr/bin/env python3

import io
import re
import requests
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_VERTICAL_ANCHOR, MSO_AUTO_SIZE, PP_PARAGRAPH_ALIGNMENT

SHORT_TEXT_LIMIT_CHARS = 75
TITLE_FONT_PT = 75
SLIDE_BLANK_LAYOUT = 6
SLIDE_WIDTH_INCHES = 10
SLIDE_HEIGHT_INCHES = 7.5
SLIDE_SMALL_MARGIN_INCHES = 0.25

debug_logs = False
debug_slide = False


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
    prs = Presentation()
    for slide in slides:
        if debug_logs:
            print("============================================ NEW SLIDE ============================================")
        layout, number_of_images = get_slide_data_layout(slide)
        fill_slide(prs, slide, layout, number_of_images)
        if debug_logs:
            print("============================================ END SLIDE ============================================")
    prs.save('test.pptx')


def get_slide_data_layout(slide):
    with_images = False
    number_of_images = 0
    with_text_blocks = False
    number_of_text_blocks = 0
    longest_text_string = 0

    for slide_data in slide:
        img_found = re.search("^img_src:(.*)$", slide_data)
        if img_found:
            if not with_images:
                with_images = True
            number_of_images += 1
        else:
            if not with_text_blocks:
                with_text_blocks = True
            number_of_text_blocks += 1
            if len(slide_data) > longest_text_string:
                longest_text_string = len(slide_data)

    if with_images and with_text_blocks:
        if number_of_images > 1 and longest_text_string < 75:
            return "images_with_columns", number_of_images
        elif number_of_images > 1:
            return "multiple_images_with_text", number_of_images
        else:
            return "one_image_with_text", number_of_images
    elif not with_images and with_text_blocks:
        if number_of_text_blocks == 1:
            return "title", number_of_images
        else:
            return "text_only", number_of_images
    elif with_images and not with_text_blocks:
        if number_of_images > 1:
            return "multiple_images", number_of_images
        else:
            return "one_image_only", number_of_images
    else:
        return "empty", number_of_images


def fill_slide(prs, slide, layout, number_of_images):
    if layout == "empty":
        return

    prs_slide_layout = prs.slide_layouts[SLIDE_BLANK_LAYOUT]
    prs_slide = prs.slides.add_slide(prs_slide_layout)

    if layout == "images_with_columns":
        slide_fill_column_layout(prs_slide, slide, number_of_images)
    else:
        generic_slide_fill(prs_slide, slide)


def generic_slide_fill(prs_slide, slide):
    images_array = []
    text_array = []

    for slide_data in slide:
        img_found = re.search("^img_src:(.*)$", slide_data)
        if img_found:
            images_array.append(img_found.group(1))
        else:
            text_array.append(slide_data)

    column_margin_inches = 0.1
    height_margin_inches = 0.1
    available_slide_width_inches = SLIDE_WIDTH_INCHES - 2*SLIDE_SMALL_MARGIN_INCHES \
        - (len(images_array)-1)*column_margin_inches
    column_width_inches = available_slide_width_inches
    if len(images_array) > 0:
        column_width_inches = available_slide_width_inches/len(images_array)
    max_image_height_inches = 0
    current_img = 0
    for img_link in images_array:
        if debug_logs:
            print("IMAGE LINK:", img_link)
        current_img += 1
        top = Inches(SLIDE_SMALL_MARGIN_INCHES)
        left = Inches(SLIDE_SMALL_MARGIN_INCHES + (current_img - 1)*column_width_inches
                      + (current_img - 1)*column_margin_inches)

        image_req = requests.get(img_link)
        img_bytes = io.BytesIO(image_req.content)
        img_box = prs_slide.shapes.add_picture(img_bytes, left, top)

        ratio = img_box.width.inches / img_box.height.inches
        if img_box.width.inches > SLIDE_WIDTH_INCHES - 2*SLIDE_SMALL_MARGIN_INCHES:
            img_box.width = Inches(SLIDE_WIDTH_INCHES - 2*SLIDE_SMALL_MARGIN_INCHES)
            img_box.height = Inches(img_box.width.inches / ratio)
        if img_box.height.inches > SLIDE_HEIGHT_INCHES - 2*SLIDE_SMALL_MARGIN_INCHES:
            img_box.height = Inches(SLIDE_HEIGHT_INCHES - 2*SLIDE_SMALL_MARGIN_INCHES)
            img_box.width = Inches(img_box.height.inches * ratio)

        if len(images_array) == 1:
            horizontal_image_center_inches = img_box.width.inches / 2
            slide_horizontal_center_inches = SLIDE_WIDTH_INCHES / 2
            left_horizontal_centered_inches = slide_horizontal_center_inches - horizontal_image_center_inches
            if left_horizontal_centered_inches < SLIDE_SMALL_MARGIN_INCHES:
                left_horizontal_centered_inches = SLIDE_SMALL_MARGIN_INCHES
            img_box.left = Inches(left_horizontal_centered_inches)
            if len(text_array) == 0:
                vertical_image_center_inches = img_box.height.inches / 2
                slide_vertical_center_inches = SLIDE_HEIGHT_INCHES / 2
                top_vertical_centered_inches = slide_vertical_center_inches - vertical_image_center_inches
                if top_vertical_centered_inches < SLIDE_SMALL_MARGIN_INCHES:
                    top_vertical_centered_inches = SLIDE_SMALL_MARGIN_INCHES
                img_box.top = Inches(top_vertical_centered_inches)

        if img_box.height.inches > max_image_height_inches:
            max_image_height_inches = img_box.height.inches

    if len(text_array) > 0:
        left = Inches(SLIDE_SMALL_MARGIN_INCHES)
        width = Inches(SLIDE_WIDTH_INCHES - 2*SLIDE_SMALL_MARGIN_INCHES)
        top = Inches(SLIDE_SMALL_MARGIN_INCHES)
        height = Inches(SLIDE_HEIGHT_INCHES - 2*SLIDE_SMALL_MARGIN_INCHES)

        if len(images_array) > 0:
            # override some default values if we have images
            top = Inches(SLIDE_SMALL_MARGIN_INCHES + max_image_height_inches + height_margin_inches)
            height = Inches(SLIDE_HEIGHT_INCHES - 2*SLIDE_SMALL_MARGIN_INCHES
                            - max_image_height_inches - height_margin_inches)

        tx_box = prs_slide.shapes.add_textbox(left, top, width, height)

        if debug_slide:
            fill = tx_box.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(255, 0, 0)

        text_frame = tx_box.text_frame
        text_frame.clear()
        text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
        text_frame.word_wrap = True
        text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

        for text in text_array:
            if debug_logs:
                print(text)
            p = text_frame.paragraphs[0]
            if p.text == "":
                if len(text_array) == 1:
                    # title-like if only one text
                    p.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                    p.font.size = Pt(TITLE_FONT_PT)
                p.text = text
            else:
                p = text_frame.add_paragraph()
                if len(text_array) == 1:
                    # title-like if only one text
                    p.alignment = PP_PARAGRAPH_ALIGNMENT.CENTER
                    p.font.size = Pt(TITLE_FONT_PT)
                p.text = text


def slide_fill_column_layout(prs_slide, slide, number_of_images):
    column_margin_inches = 0.1
    height_margin_inches = 0.1
    available_slide_width_inches = SLIDE_WIDTH_INCHES - 2*SLIDE_SMALL_MARGIN_INCHES \
        - (number_of_images-1)*column_margin_inches
    column_width_inches = available_slide_width_inches/number_of_images
    current_img = 0
    new_image = True

    for slide_data in slide:
        width = Inches(column_width_inches)
        img_found = re.search("^img_src:(.*)$", slide_data)
        if img_found:
            if debug_logs:
                print("IMAGE LINK:", img_found.group(1))
            new_image = True
            current_img += 1
            top = Inches(SLIDE_SMALL_MARGIN_INCHES)
            left = Inches(SLIDE_SMALL_MARGIN_INCHES + (current_img - 1)*column_width_inches
                          + (current_img - 1)*column_margin_inches)

            image_req = requests.get(img_found.group(1))
            img_bytes = io.BytesIO(image_req.content)
            img_box = prs_slide.shapes.add_picture(img_bytes, left, top)
            current_img_height_inches = img_box.height.inches

            # with open("test.jpg", 'wb') as out:
            #     out.write(img_bytes.read())
        else:
            if debug_logs:
                print(slide_data)
            if new_image:
                new_image = False
                left = Inches(SLIDE_SMALL_MARGIN_INCHES + (current_img - 1)*column_width_inches
                              + (current_img - 1)*column_margin_inches)
                top = Inches(SLIDE_SMALL_MARGIN_INCHES + current_img_height_inches + height_margin_inches)
                height = Inches(SLIDE_HEIGHT_INCHES - 2*SLIDE_SMALL_MARGIN_INCHES
                                - current_img_height_inches - height_margin_inches)

                tx_box = prs_slide.shapes.add_textbox(left, top, width, height)

                if debug_slide:
                    fill = tx_box.fill
                    fill.solid()
                    fill.fore_color.rgb = RGBColor(255, 0, 0)

                text_frame = tx_box.text_frame
                text_frame.clear()
                text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.TOP
                text_frame.word_wrap = True
                text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

                p = text_frame.paragraphs[0]
                p.text = slide_data
            else:
                p = text_frame.add_paragraph()
                p.text = slide_data


# html_to_pptx("", "")
