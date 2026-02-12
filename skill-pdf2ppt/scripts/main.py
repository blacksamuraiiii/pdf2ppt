
import argparse
import os
import sys
import logging

# Ensure project root is in sys.path
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

try:
    from core.converter import convert_pdf_to_ppt
except ImportError:
    # If core is not found, try adding parent dir
    sys.path.append(os.path.dirname(current_dir))
    from core.converter import convert_pdf_to_ppt

def setup_logging(debug=False):
    level = logging.DEBUG if debug else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout)
        ]
    )

def main():
    parser = argparse.ArgumentParser(description="Convert PDF to PowerPoint (PPTX)")
    
    parser.add_argument("--pdf_path", required=True, help="Path to the input PDF file")
    parser.add_argument("--token", help="MinerU API Token (can also be set via MINERU_TOKEN env var)")
    parser.add_argument("--output_path", help="Path to the output PPTX file")
    parser.add_argument("--ratio", choices=["16:9", "4:3"], default="16:9", help="Aspect ratio of the slides (default: 16:9)")
    parser.add_argument("--remove_watermark", action="store_true", help="Remove watermarks from the PDF")
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")
    
    args = parser.parse_args()
    
    setup_logging(args.debug)
    
    # 1. Check Token
    token = args.token or os.environ.get("MINERU_TOKEN")
    if not token:
        logging.error("Error: MinerU Token is required. Provide it via --token or MINERU_TOKEN environment variable.")
        sys.exit(1)
    
    # 2. Check Input File
    if not os.path.exists(args.pdf_path):
        logging.error(f"Error: PDF file not found at {args.pdf_path}")
        sys.exit(1)
        
    # 3. Determine Output Path
    if args.output_path:
        ppt_path = args.output_path
    else:
        base_name = os.path.splitext(args.pdf_path)[0]
        ppt_path = base_name + ".pptx"
    
    # 4. Determine Aspect Ratio
    if args.ratio == "4:3":
        w, h = 10, 7.5
    else:
        w, h = 16, 9
        
    # 5. Execute Conversion
    try:
        logging.info(f"Starting conversion: {args.pdf_path} -> {ppt_path}")
        convert_pdf_to_ppt(
            pdf_input_path=args.pdf_path,
            ppt_output_path=ppt_path,
            mineru_token=token,
            ppt_slide_width=w,
            ppt_slide_height=h,
            use_cache=True, # Default to using cache for CLI
            cache_dir=os.path.join(os.path.dirname(ppt_path), "temp"),
            remove_watermark=args.remove_watermark
        )
        logging.info("Conversion Success")
    except Exception as e:
        logging.error(f"Conversion Failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
