# -*- coding: utf-8 -*-


import zipfile

from pandas import DataFrame

from bs4 import BeautifulSoup


def extract_comments(path, comments_file='word/comments.xml'):
    """Return the comments from the `.docx` file in `path` as a DataFrame.

    Parameters
    ==========
    path : string
        Location of Word document
    comments_file : string
        Location of comments file inside `path`

    Returns
    =======
    comments : pandas DataFrame
        Comments from `path` as a table

    """
    # Extract comments file from `path`
    try:
        with zipfile.ZipFile(path) as z:
            if comments_file not in z.namelist():
                comments = None
            else:
                comments = z.open(comments_file).read()
    except:
        raise ValueError('Unable to open document: %s' % path)
    # Process comments
    if comments is not None:
        soup = BeautifulSoup(comments)
        contents = soup.body.contents[0]
        comments = {}
        for chunk in contents:
            # Extract attributes and rename
            comment = chunk.attrs
            comment = {k.replace('w:', ''): v for k, v in comment.items()}
            # Parse separate date and time
            date = comment['date']
            date, time = date.split('T')
            comment['date'] = date
            comment['time'] = time.replace('Z', '')
            # Extract comment text
            text = chunk.find_all('w:t')
            text = [t.get_text() for t in text]
            text = ''.join(text)
            comment['text'] = text
            # Store with 'id' as the key (used in DataFrame index)
            comments[comment['id']] = comment
        comments = DataFrame(comments).transpose()
    # Reorder
    comments = comments[['initials', 'author', 'date', 'time', 'text']]
    # Return
    return comments
