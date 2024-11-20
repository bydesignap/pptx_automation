from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

def add_widget_two_tone(
    slide,
    width_inches=3.87,
    top_height_inches=0.51,
    bottom_height_inches=0.75,
    position_x_inches=None,  # New parameter for horizontal positioning
    position_y_inches=2,
    title_text="Cell Title",
    left_text_lines=["3Q24", "3Q24 YTD"],
    right_text_lines=["$12.70 billion", "$39.64 billion"],
    title_font_size=14,
    body_font_size=12,
    border_width_pt=1,
    primary_color=RGBColor(31, 57, 108),  # #1e3a8a
    background_color=RGBColor(255, 255, 255)  # white
):
    """
    Creates a PowerPoint slide with revenue information in a custom shape design.
    
    Parameters:
    -----------
    width_inches : float
        Width of the shapes in inches
    top_height_inches : float
        Height of the top shape in inches
    bottom_height_inches : float
        Height of the bottom shape in inches
    position_x_inches : float or None
        Horizontal position from left of slide in inches. If None, centers horizontally
    position_y_inches : float
        Vertical position from top of slide in inches
    title_text : str
        Text for the title
    left_text_lines : list
        List of strings for the left column
    right_text_lines : list
        List of strings for the right column
    title_font_size : int
        Font size for the title text
    body_font_size : int
        Font size for the body text
    border_width_pt : int
        Width of the border in points
    primary_color : RGBColor
        Primary color for fills and borders
    background_color : RGBColor
        Background color for bottom shape
    """
    
    
    # Convert dimensions to PowerPoint units
    width = Inches(width_inches)
    top_height = Inches(top_height_inches)
    bottom_height = Inches(bottom_height_inches)
    
    # Calculate horizontal position
    if position_x_inches is None:
        # Center horizontally if no position specified
        left = (prs.slide_width - width) / 2
    else:
        left = Inches(position_x_inches)
    
    top_y = Inches(position_y_inches)
    
    # Add top rounded rectangle
    top_shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUND_2_SAME_RECTANGLE,
        left,
        top_y,
        width,
        top_height
    )
    
    # Format top shape
    top_shape.fill.solid()
    top_shape.fill.fore_color.rgb = primary_color
    top_shape.line.width = Pt(border_width_pt)
    top_shape.line.color.rgb = primary_color
    top_shape.shadow.inherit = False
    
    # Add bottom rounded rectangle (rotated 180°)
    bottom_shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUND_2_SAME_RECTANGLE,
        left,
        top_y + top_height,
        width,
        bottom_height
    )
    
    # Format bottom shape
    bottom_shape.fill.solid()
    bottom_shape.fill.fore_color.rgb = background_color
    bottom_shape.line.width = Pt(border_width_pt)
    bottom_shape.line.color.rgb = primary_color
    bottom_shape.shadow.inherit = False
    bottom_shape.rotation = 180
    
    # Add title text
    title = slide.shapes.add_textbox(
        left,
        top_y + Inches(0.1),
        width,
        top_height
    )
    title_frame = title.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.alignment = PP_ALIGN.CENTER
    title_run = title_para.add_run()
    title_run.text = title_text
    title_run.font.size = Pt(title_font_size)
    title_run.font.bold = True
    title_run.font.color.rgb = background_color
    
    # Add left column text
    left_text = slide.shapes.add_textbox(
        left + Inches(0.2),
        top_y + top_height + Inches(0.1),
        Inches(1),
        bottom_height - Inches(0.2)
    )
    left_frame = left_text.text_frame
    
    # Add all left column lines
    for i, text in enumerate(left_text_lines):
        if i == 0:
            p = left_frame.paragraphs[0]
        else:
            p = left_frame.add_paragraph()
            p.space_before = Pt(6)
        
        run = p.add_run()
        run.text = text
        run.font.size = Pt(body_font_size)
        run.font.bold = True
        run.font.color.rgb = primary_color
    
    # Add right column text
    right_text = slide.shapes.add_textbox(
        left + Inches(1.2),
        top_y + top_height + Inches(0.1),
        width - Inches(1.4),
        bottom_height - Inches(0.2)
    )
    right_frame = right_text.text_frame
    
    # Add all right column lines
    for i, text in enumerate(right_text_lines):
        if i == 0:
            p = right_frame.paragraphs[0]
        else:
            p = right_frame.add_paragraph()
            p.space_before = Pt(6)
        
        p.alignment = PP_ALIGN.RIGHT
        run = p.add_run()
        run.text = text
        run.font.size = Pt(body_font_size)
        run.font.bold = True
        run.font.color.rgb = primary_color
    
    return slide

def add_centered_line(slide, line_x=Inches(.47), line_y=Inches(4.03), line_width=Inches(12.52), line_weight=Pt(4)):
    # Calculate line position
    
    # Add line shape
    line = slide.shapes.add_shape(
        MSO_SHAPE.LINE_INVERSE,
        line_x, line_y,
        line_width, 0
    )
    
    # Format line
    line.line.color.rgb = RGBColor(31, 57, 108)
    line.line.width = line_weight # 4 pixels wide

def add_content_box(prs, slide):
    # Get slide dimensions
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    
    # Add gray box
    box_width = Inches(13.33)
    box_height = Inches(6.49)
    # Calculate left position to center the box
    box_left = (slide_width - box_width) / 2
    # Calculate top position so bottom aligns with slide bottom
    box_top = slide_height - box_height
    
    box = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        box_left,
        box_top,
        box_width,
        box_height
    )
    
    # Format box
    fill = box.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(227, 228, 231)  # Gray color
    # Make the box border transparent
    box.line.color.rgb = RGBColor(227, 228, 231)

if __name__ == '__main__':
    # Create presentation and slide
    prs = Presentation('template.pptx')
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[3])  # blank layout
    slide.shapes.title.text = "Risk Summary"

    # Add content box
    add_content_box(prs, slide)

    # Add centered line
    add_centered_line(slide)

    # Example usage with default parameters (centered)
    add_widget_two_tone(slide, position_x_inches=.47, position_y_inches=1.24, title_text="Technology Operations", left_text_lines=["3Q24", "4Q24"], right_text_lines=["YELLOW", "GREEN"])
    add_widget_two_tone(slide, position_x_inches=4.73, position_y_inches=1.24, title_text="Technology Development", left_text_lines=["3Q24", "4Q24"], right_text_lines=["YELLOW", "GREEN"])
    add_widget_two_tone(slide, position_x_inches=9.12, position_y_inches=1.24, title_text="Technology Resiliency", left_text_lines=["3Q24", "4Q24"], right_text_lines=["YELLOW", "GREEN"])

    add_widget_two_tone(slide, position_x_inches=.47, position_y_inches=2.56, title_text="Information & Asset Management", left_text_lines=["3Q24", "4Q24"], right_text_lines=["YELLOW", "GREEN"])
    add_widget_two_tone(slide, position_x_inches=4.73, position_y_inches=2.56, title_text="Technology Stability", left_text_lines=["3Q24", "4Q24"], right_text_lines=["YELLOW", "GREEN"])
    add_widget_two_tone(slide, position_x_inches=9.12, position_y_inches=2.56, title_text="Technology Modenization", left_text_lines=["3Q24", "4Q24"], right_text_lines=["YELLOW", "GREEN"])

    # Save the presentation
    prs.save('risk_summary_slide.pptx')



    # Example usage with specific horizontal position
    """
    create_revenue_slide(
        width_inches=3.87,
        position_x_inches=2.0,  # Position 2 inches from left
        position_y_inches=1.5,
        title_text="Annual Revenue",
        left_text_lines=["2023", "2024 (YTD)"],
        right_text_lines=["$45.2B", "$32.1B"]
    )
    """