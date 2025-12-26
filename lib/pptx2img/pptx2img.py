import os
import sys
import ctypes

# Try to import win32com, prompt user if missing
try:
    import win32com.client
except ImportError:
    print("Error: Library 'pywin32' is required. Please install it via: pip install pywin32")
    sys.exit(1)

def topng(pptx, output_dir="./pptx2img", range=None, scale=None):
    """
    Convert PowerPoint slides to PNG images.

    Args:
        pptx (str): Path to the .pptx file.
        output_dir (str): Directory to save the images.
        range (list): Optional. A list [start, end] specifying slide range (1-based).
                      Example: [1, 5] converts slides 1 to 5.
        scale (int): Optional. Resolution scale.
                     If None or 0, it adapts to the screen's long edge resolution.
                     If specified (e.g., 1, 2), it scales relative to original slide points.
    """
    # 1. Path handling
    pptx_path = os.path.abspath(pptx)
    output_path = os.path.abspath(output_dir)

    if not os.path.exists(pptx_path):
        print("Error: File '%s' not found." % pptx_path)
        return

    if not os.path.exists(output_path):
        os.makedirs(output_path)
        print("Created output directory: %s" % output_path)

    # 2. Initialize PowerPoint Application
    powerpoint = None
    presentation = None
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    except Exception as e:
        print("Error: Could not initialize PowerPoint. Make sure Microsoft PowerPoint is installed.")
        print("Details: %s" % e)
        return

    try:
        # 3. Open Presentation (WithWindow=False attempts background processing)
        presentation = powerpoint.Presentations.Open(pptx_path, WithWindow=False)

        # 4. Determine Slide Range
        total_slides = presentation.Slides.Count
        start_slide = 1
        end_slide = total_slides

        if range and isinstance(range, list) and len(range) == 2:
            start_slide = max(1, range[0])
            end_slide = min(total_slides, range[1])

        # 5. Calculate Target Resolution
        slide_width = presentation.PageSetup.SlideWidth
        slide_height = presentation.PageSetup.SlideHeight
        
        target_w = 0
        target_h = 0

        # Logic: If scale is not provided, use screen resolution (Long Edge)
        if not scale:
            try:
                user32 = ctypes.windll.user32
                screen_w = user32.GetSystemMetrics(0)
                screen_h = user32.GetSystemMetrics(1)
                
                # Find the long edge of the screen
                screen_long = max(screen_w, screen_h)
                
                # Calculate aspect ratio of the slide
                slide_ratio = slide_width / slide_height

                if slide_ratio >= 1: # Landscape Slide
                    target_w = screen_long
                    target_h = int(screen_long / slide_ratio)
                else: # Portrait Slide
                    target_h = screen_long
                    target_w = int(screen_long * slide_ratio)
                
                print("Mode: Auto-Resolution (Matched to Screen Long Edge: %d px)" % screen_long)
            except Exception:
                # Fallback if ctypes fails
                target_w = int(slide_width * 2)
                target_h = int(slide_height * 2)
                print("Mode: Fallback Resolution (2x)")
        else:
            # Manual scale
            target_w = int(slide_width * scale)
            target_h = int(slide_height * scale)
            print("Mode: Manual Scale (%dx)" % scale)

        print("Processing '%s'..." % os.path.basename(pptx))
        print("Target Size: %dx%d px" % (target_w, target_h))
        print("Converting slides %d to %d..." % (start_slide, end_slide))

        # 6. Iterate and Export
        count = 0
        for i in range(start_slide, end_slide + 1):
            slide = presentation.Slides(i)
            # Filename format: Slide_1.png, Slide_2.png
            image_name = "Slide_%d.png" % i
            image_path = os.path.join(output_path, image_name)

            # Export to PNG
            slide.Export(image_path, "PNG", target_w, target_h)
            count += 1
            print("Saved: %s" % image_name)

        print("Done! %d images saved to '%s'." % (count, output_path))

    except Exception as e:
        print("An error occurred during conversion: %s" % e)
    finally:
        # 7. Cleanup Resources
        if presentation:
            presentation.Close()
        # powerpoint.Quit() is optional; typically omitted to avoid killing user's other open PPTs
        pass

def whatis():
    """Prints the library information."""
    info = """
--------------------------------------------------
pptx2img Info
--------------------------------------------------
Version     : 2025v1
Author      : WaterRun
GitHub      : https://github.com/Water-Run/pptx2img
Email       : 2263633954@qq.com
Description : A library to convert PPTX to PNG.

Note: A Windows GUI (EXE) version is also available:
https://github.com/Water-Run/pptx2img/releases/tag/pptx2img
--------------------------------------------------
    """
    print(info.strip())