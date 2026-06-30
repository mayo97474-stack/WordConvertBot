import os
import sys
import logging
import io
from datetime import datetime
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# For Word to PDF conversion
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import tempfile
import subprocess

# Enable logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

# Get bot token from environment variable
def get_token():
    """Get bot token from environment variables."""
    token = os.environ.get('BOT_TOKEN')
    if not token:
        token = os.environ.get('TELEGRAM_BOT_TOKEN')
    if not token:
        logger.error("❌ No BOT_TOKEN found in environment variables!")
        logger.error("Please add BOT_TOKEN to your Railway Variables.")
        sys.exit(1)
    return token

TOKEN = get_token()
logger.info("✅ Bot token loaded successfully!")

# Store user sessions
user_sessions = {}

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send a welcome message."""
    user = update.effective_user
    welcome_text = f"""
📄 **Welcome to WordConvertBot, {user.first_name}!**

I convert Word documents (.docx) to PDF format.

**How to use:**
1️⃣ Send me a .docx file
2️⃣ I'll convert it to PDF
3️⃣ Get your PDF file back!

**Commands:**
/start - Show this welcome message
/help - Show all commands
/convert - Convert a Word file (send as document)

**Example:**
Just send me any Word document (.docx) and I'll convert it!

💡 **Tip:** Works with .docx files only.
"""
    await update.message.reply_text(welcome_text)


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send a help message."""
    help_text = """
📖 **How to use WordConvertBot:**

1️⃣ Send me a .docx file
2️⃣ I'll convert it to PDF
3️⃣ Get your PDF file back!

**Commands:**
/start - Welcome message
/help - Show this help message
/convert - Convert a Word file

**Supported formats:**
• .docx → .pdf

**What you can convert:**
• Word documents
• Reports
• Resumes/CVs
• Any .docx file!

💡 **Pro tip:** Your file is processed securely and deleted after conversion.
"""
    await update.message.reply_text(help_text)


async def convert_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle /convert command."""
    await update.message.reply_text(
        "📤 Please send me a .docx file to convert!\n\n"
        "Just attach the file and I'll convert it to PDF."
    )


async def convert_to_pdf(docx_data: bytes) -> bytes:
    """
    Convert .docx file to PDF using python-docx and ReportLab.
    This is a simple implementation that preserves text formatting.
    """
    try:
        # Create temporary files
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_docx:
            temp_docx.write(docx_data)
            docx_path = temp_docx.name
        
        # Read the Word document
        doc = Document(docx_path)
        
        # Create PDF using ReportLab
        pdf_buffer = io.BytesIO()
        c = canvas.Canvas(pdf_buffer, pagesize=A4)
        width, height = A4
        
        # Set up text formatting
        y = height - inch
        left_margin = inch
        font_size = 11
        
        # Process each paragraph
        for paragraph in doc.paragraphs:
            if paragraph.text.strip():
                # Check if it's a heading (styled)
                if paragraph.style.name.startswith('Heading'):
                    c.setFont('Helvetica-Bold', font_size + 2)
                else:
                    c.setFont('Helvetica', font_size)
                
                # Handle text wrapping
                text = paragraph.text
                lines = []
                current_line = ""
                for word in text.split():
                    if len(current_line + word + " ") * 5 < (width - 2 * left_margin):
                        current_line += word + " "
                    else:
                        if current_line:
                            lines.append(current_line.strip())
                        current_line = word + " "
                if current_line:
                    lines.append(current_line.strip())
                
                # Draw each line
                for line in lines:
                    if y < inch:
                        c.showPage()
                        c.setFont('Helvetica', font_size)
                        y = height - inch
                    c.drawString(left_margin, y, line)
                    y -= (font_size + 2)
                
                # Add spacing between paragraphs
                y -= 4
        
        c.save()
        pdf_buffer.seek(0)
        
        # Clean up temp file
        os.unlink(docx_path)
        
        return pdf_buffer.getvalue()
        
    except Exception as e:
        logger.error(f"PDF conversion error: {e}")
        raise e


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle document files sent to the bot."""
    document = update.message.document
    
    # Check if it's a .docx file
    if document.file_name and document.file_name.lower().endswith('.docx'):
        # Send processing message
        processing_msg = await update.message.reply_text(
            "🔄 **Converting your Word document to PDF...**\n\n"
            f"📄 File: `{document.file_name}`\n"
            f"⏳ This may take a moment..."
        )
        
        try:
            # Download the file
            file = await document.get_file()
            file_data = await file.download_as_bytearray()
            
            # Convert to PDF
            pdf_data = await convert_to_pdf(file_data)
            
            # Generate new filename
            base_name = document.file_name.rsplit('.', 1)[0]
            pdf_filename = f"{base_name}.pdf"
            
            # Send the PDF back
            await update.message.reply_document(
                document=io.BytesIO(pdf_data),
                filename=pdf_filename,
                caption=f"✅ **Converted successfully!**\n\n"
                        f"📄 Original: `{document.file_name}`\n"
                        f"📄 Converted: `{pdf_filename}`\n"
                        f"📊 Size: {len(pdf_data) / 1024:.1f} KB"
            )
            
            await processing_msg.delete()
            
        except Exception as e:
            logger.error(f"Conversion failed: {e}")
            await processing_msg.edit_text(
                f"❌ **Sorry, I couldn't convert that file.**\n\n"
                f"Error: {str(e)}\n\n"
                f"💡 **Tips:**\n"
                f"• Make sure it's a valid .docx file\n"
                f"• Try a different file\n"
                f"• Use /help for guidance"
            )
    else:
        await update.message.reply_text(
            "❌ **Please send a .docx file!**\n\n"
            "I can only convert Word documents (.docx) to PDF.\n"
            f"File type: `{document.file_name}`"
        )


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle text messages that aren't commands."""
    await update.message.reply_text(
        "📤 Please send me a .docx file to convert.\n\n"
        "Use /help for more information."
    )


def main() -> None:
    """Start the bot."""
    try:
        # Create Application
        application = Application.builder().token(TOKEN).build()
        
        # Add command handlers
        application.add_handler(CommandHandler("start", start))
        application.add_handler(CommandHandler("help", help_command))
        application.add_handler(CommandHandler("convert", convert_command))
        
        # Add message handlers
        application.add_handler(MessageHandler(filters.Document.ALL, handle_document))
        application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
        
        # Start the Bot
        logger.info("🚀 WordConvertBot started successfully!")
        logger.info("📄 Press Ctrl+C to stop.")
        application.run_polling()
        
    except Exception as e:
        logger.error(f"❌ Fatal error: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
