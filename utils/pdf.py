from reportlab.pdfgen import canvas as can
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# register custom colour
pdfmetrics.registerFont(TTFont('OpenSans', './fonts/OpenSans-Light.ttf'))


def create_cert(name):
    """Given student's name, generate certificate of completion

    :param name: Student's name from google sheets
    """
    canvas = can.Canvas('aaa.pdf', pagesize=landscape(letter))

    page_width, page_height = canvas._pagesize

    canvas.drawImage('./img.jpg', 0, 0, width=page_width, height=page_height,
                     mask=None, preserveAspectRatio=True)

    # for some reason reportlab uses rbg colors as ratios of 256
    col = 102.0/256
    canvas.setFillColorRGB(col, col, col)

    canvas.setFont('OpenSans', 24)
    canvas.drawCentredString(400, (page_height/2)+20, name)

    canvas.showPage()
    canvas.save()
