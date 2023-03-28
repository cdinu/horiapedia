# Horiapedia

We have a folder with about 1000 powerpoint PPTX files. It is a knowlege database organised like this:
- there are categories and subcategories
- subcategories can have sub-subcategories, which in turn can have sub-sub-subcategories. the depth is arbitrary, but less than 8
- the reference for a piece of content is category.subcategory.sub-subcategory.sub-sub-sub-category.etc.final-subcategory.
- the powepoint file titles are reference.pptx, for example b.g.pop.coalescence.pptx or u.museum.british.apollonian.ppt and so on. for the last one the category is b, the subcategory is museum, the sub-subcategory is british, and the final-subcategory is apollonian.
- inside the powerpoint there are slides. some slides contain one or more texts in the form of 2011_08_18, which is a date in the format YYYY_MM_DD. the first date is the date of creation, the rest are the dates for updates.
- inside the powepoint file there are two types of slides:
    - "reference" slides, which are the slides that contain as the first text a reference in the form of cat.subcat.subsubcat.etc.
    - "original" or "content" slides, which do not have the reference as the first text.

This software will eventually do the following:
- take as input the folder with the powerpoint files
- for each pptx file, 
    - the working directory is category/subcategory/sub-subcategory/sub-sub-sub-category/etc/final-subcategory. create it if it does not exist
    - reads the slides in the powerpoint and for each slide:
        - read the title, the creation and update dates, the reference and the content (text and images)
        - if the slide is "original"
            - create a markdown file with the name <title>.md and the following content:
                - yaml front matter with the following fields:
                    - title
                    - creation date
                    - update dates
                    - reference
                    - content
                - heading 1 with title
                - creation date
                - update dates
                - reference
                - content (text and images)
            - images should be extracted from the powerpoint and saved in the same folder as the markdown file, named <title>-<image number>.png
        - if the slide is "reference", create a markdown file with the name <title>.md and the following content:
            - create the markdown file with the same content
                - it should have the following content at the beginning:
                    - "This slide is origianlly found in <reference to the other subcategory>"
                    - the reference is also a link to the referred content (as folder)
                    - the link should be relative to the current folder
                    - the link should be to the folder, not to the markdown file
                - the name of the markdown file is "ref-<title>.md"
    - generate a markdown file with the name "index.md" and the following content:
        - yaml front matter with the following fields:
            - title
            - creation date
            - update dates
            - reference
            - content
        - heading 1 with title
        - creation date
        - a table containing the list of all the titles, in order of the slides
