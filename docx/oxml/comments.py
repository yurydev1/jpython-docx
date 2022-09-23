"""
Custom element classes related to the comments part
"""

from . import OxmlElement
from .simpletypes import ST_DecimalNumber, ST_String
from ..opc.constants import NAMESPACE
from ..text.paragraph import Paragraph
from ..text.run import Run
from .xmlchemy import (
    BaseOxmlElement, OneAndOnlyOne, RequiredAttribute, ZeroOrMore, ZeroOrOne
)

# files:
#  - word/settings.xml
#     - rsids:        Specifies a unique identifier used to track the editing session when the section mark for this section was added to the document.
#                     All rsid* attributes throughout this document of an equal value, if present, shall indicate that those regions were modified during the same editing session.
#                     A producer may choose to increment the revision save ID value to indicate subsequent editing sessions to indicate the order of the saves performed.
#  - word/commentsIds.xml, https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/9c360cd7-653f-4d82-82be-7bda2488c0c1
#     - durableId:    specifies the identifier for the associated comment
#     - paraId:       specifies the paraId (section 2.6.2.3) of the last paragraph of the associated comment
#  - word/commentsExtended.xml
#     - done:         specifies whether the comment has been addressed (0 or 1)
#     - paraId:       specifies the paraId (section 2.6.2.3) of the last paragraph of the associated comment
#  - word/commentsExtensible.xml
#     - dateUtc:      some dateUtc
#     - durableId:    specifies the identifier for the associated comment

class CT_Com(BaseOxmlElement):
    """
    A ``<w:comment>`` element, a container for Comment properties
    """
    initials = RequiredAttribute('w:initials', ST_String)
    _id = RequiredAttribute('w:id', ST_DecimalNumber)
    date = RequiredAttribute('w:date', ST_String)
    author = RequiredAttribute('w:author', ST_String)

    p = ZeroOrMore('w:p', successors=('w:comment',))

    @classmethod
    def new(cls, initials, comm_id, date, author):
        """
        Return a new ``<w:comment>`` element having _id of *comm_id* and having
        the passed params as meta data
        """
        comment = OxmlElement('w:comment')
        comment.initials = initials
        comment.date = date
        comment._id = comm_id
        comment.author = author
        return comment

    def _add_p(self, text):
        _p = OxmlElement('w:p')
        _r = _p.add_r()
        run = Run(_r,self)
        run.text = text
        self._insert_p(_p)
        return _p

    @property
    def meta(self):
        return [self.author, self.initials, self.date]

    @property
    def paragraphs(self):
        return [Paragraph(p, self) for p in self.p_lst]


class CT_Comments(BaseOxmlElement):
    """
    A ``<w:comments>`` element, a container for Comments properties
    """
    comment = ZeroOrMore ('w:comment', successors=('w:comments',))

    def add_comment(self,author, initials, date):
        _next_id = self._next_commentId
        comment = CT_Com.new(initials, _next_id, date, author)
        comment = self._insert_comment(comment)

        return comment

    @property
    def _next_commentId(self):
        ids = self.xpath('./w:comment/@w:id')
        len(ids)
        _ids = [int(_str) for _str in ids]
        _ids.sort()

        try:
            return _ids[-1] + 2
        except:
            return 0

    def get_comment_by_id(self, _id):
        namesapce = NAMESPACE().WML_MAIN
        for c in self.findall('.//w:comment',{'w':namesapce}):
            if c._id == _id:
                return c
        return None


class CT_CRS(BaseOxmlElement):
    """
    A ``<w:commentRangeStart>`` element
    """
    _id = RequiredAttribute('w:id', ST_DecimalNumber)

    @classmethod
    def new(cls, _id):
        commentRangeStart = OxmlElement('w:commentRangeStart')
        commentRangeStart._id =_id

        return commentRangeStart

class CT_CRE(BaseOxmlElement):
    """
    A ``w:commentRangeEnd`` element
    """
    _id = RequiredAttribute('w:id', ST_DecimalNumber)


    @classmethod
    def new(cls, _id):
        commentRangeEnd = OxmlElement('w:commentRangeEnd')
        commentRangeEnd._id =_id
        return commentRangeEnd


class CT_CRef(BaseOxmlElement):
    """
    w:commentReference
    """
    _id = RequiredAttribute('w:id', ST_DecimalNumber)

    @classmethod
    def new (cls, _id):
        commentReference = OxmlElement('w:commentReference')
        commentReference._id =_id
        return commentReference


