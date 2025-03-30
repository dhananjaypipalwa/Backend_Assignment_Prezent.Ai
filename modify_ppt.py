import aspose.slides as slides
from aspose.pydrawing import Color

# ========================== Slide 1: Modify Image & Labels ========================== #
def modify_slide1(slide, presentation):
    """Replace image with a car image and update labels from 01-06 to A-F."""
    try:
        with open("car_image.jpg.jpg", "rb") as img_file:
            car_image = presentation.images.add_image(img_file)

        for shape in slide.shapes:
            if isinstance(shape, slides.PictureFrame):
                shape.picture_format.picture.image = car_image

        labels = ["A", "B", "C", "D", "E", "F"]
        for shape in slide.shapes:
            if isinstance(shape, slides.AutoShape) and shape.text_frame is not None:
                text = shape.text_frame.text.strip()
                if text in ["01", "02", "03", "04", "05", "06"]:
                    shape.text_frame.text = labels[int(text) - 1]  # Convert numbers to A-F
    except Exception as e:
        print(f"❌ Error modifying Slide 1: {e}")

# ========================== Slide 2: Modify Table Colors ========================== #
def modify_slide2(slide):
    """Change the header color of columns 2-7 to match column 1."""
    try:
        for shape in slide.shapes:
            if isinstance(shape, slides.Table):
                first_column_color = shape.rows[0][0].cell_format.fill_format.solid_fill_color.color
                for col in range(1, 7):  # Modify columns 2-7
                    shape.rows[0][col].cell_format.fill_format.solid_fill_color.color = first_column_color
    except Exception as e:
        print(f"❌ Error modifying Slide 2: {e}")

# ========================== Slide 3: Modify Title & Chart Colors ========================== #
def modify_slide3(slide):
    """Change the title color to blue and update chart colors to red."""
    try:
        for shape in slide.shapes:
            if isinstance(shape, slides.AutoShape) and shape.text_frame is not None:
                portion = shape.text_frame.paragraphs[0].portions[0]
                portion.portion_format.fill_format.solid_fill_color.color = Color.blue  # Change title color

        for shape in slide.shapes:
            if isinstance(shape, slides.charts.Chart):  # Corrected `charts.Chart`
                for series in shape.chart_data.series:
                    for point in series.data_points:
                        point.format.fill.fill_type = slides.FillType.SOLID
                        point.format.fill.solid_fill_color.color = Color.red  # Change chart color to red
    except Exception as e:
        print(f"❌ Error modifying Slide 3: {e}")

# ========================== Main Function to Modify PPT ========================== #
def modify_ppt(input_ppt, output_ppt):
    """Modify the PowerPoint file based on requirements."""
    try:
        with slides.Presentation(input_ppt) as presentation:
            if len(presentation.slides) < 3:
                print("❌ PowerPoint file must have at least 3 slides.")
                return

            print("🔹 Modifying Slide 1...")
            modify_slide1(presentation.slides[0], presentation)

            print("🔹 Modifying Slide 2...")
            modify_slide2(presentation.slides[1])

            print("🔹 Modifying Slide 3...")
            modify_slide3(presentation.slides[2])

            # Save the modified PowerPoint file
            presentation.save(output_ppt, slides.export.SaveFormat.PPTX)
            print(f"✅ Modified PowerPoint saved as '{output_ppt}'")

    except Exception as e:
        print(f"❌ Error modifying PowerPoint: {e}")

# ========================== Run the Script ========================== #
if __name__ == "__main__":
    input_ppt = "assignment-1.pptx"
    output_ppt = "modified_assignment.pptx"
    modify_ppt(input_ppt, output_ppt)
