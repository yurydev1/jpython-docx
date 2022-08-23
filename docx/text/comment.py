from ..shared import Parented

class Comment(Parented):
    """[summary]

    :param Parented: [description]
    :type Parented: [type]
    """
    def __init__(self, com, parent):
        super(Comment, self).__init__(parent)
        self._com = self._element = self.element = com
    
    @property
    def paragraphs(self):
        return self.element.paragraphs
    
    @property
    def text(self):
        return '\n'.join([p.text for p in self.element.paragraphs])
    
    @text.setter
    def text(self, text):
        self.element.paragraphs.text = text