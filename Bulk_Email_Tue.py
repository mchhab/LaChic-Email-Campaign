import smtplib
import ssl
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import time
import os
import hashlib
import urllib.parse
from io import BytesIO
from PIL import Image  # for resizing/compressing images

# ==========================
# CONFIGURATION
# ==========================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_PATH = os.path.join(BASE_DIR, "data", "Bulk_Email_Tue.xlsx")
EMAIL_COLUMN = "Email"
NAME_COLUMN = "Name"

SMTP_SERVER = "smtp.ionos.com"
SMTP_PORT = 587
EMAIL_ADDRESS = "manik@lachicdesigns.com"
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
if not EMAIL_PASSWORD:
    raise RuntimeError("EMAIL_PASSWORD env var not set")

HANDBAG_IMAGE_PATH_1 = os.path.join(BASE_DIR, "images", "LaChic-090.jpg")
HANDBAG_IMAGE_PATH_2 = os.path.join(BASE_DIR, "images", "LaChic-094.jpg")

EMAIL_SUBJECT = "✨ Re: Following up on your interest in our Handbags - La Chic Designs!"

BASE_TRACKING_URL = "https://lachicdesigns.com/"
OPEN_PIXEL_BASE = "https://lachicdesigns.com/email-open"

UTM_SOURCE   = "email"
UTM_MEDIUM   = "bulk"
UTM_CAMPAIGN = "handbag_drop_2025_01"

CATALOG_URL = "https://drive.google.com/file/d/1IvZvhlwjSTF0e8HQ7HJZwa74fk1HVyk1/view?usp=share_link"


# ==========================
# TRACKING HELPERS
# ==========================

def build_tracking(email):
    """Build tracking URL + open pixel."""
    email_hash = hashlib.md5(email.lower().encode("utf-8")).hexdigest()

    params = {
        "utm_source": UTM_SOURCE,
        "utm_medium": UTM_MEDIUM,
        "utm_campaign": UTM_CAMPAIGN,
        "utm_content": email_hash,
    }

    tracking_url = BASE_TRACKING_URL.rstrip("/") + "/?" + urllib.parse.urlencode(params)
    pixel_url = f"{OPEN_PIXEL_BASE}?id={email_hash}"
    return tracking_url, pixel_url


# ==========================
# IMAGE HELPER (resize/compress)
# ==========================

def make_inline_image(path: str, cid: str, max_size=(800, 800), quality=70):
    """
    Load an image, resize it to fit within max_size, compress it, and
    return a MIMEImage with the given Content-ID. Returns None if file missing.
    """
    if not os.path.isfile(path):
        print(f"Warning: image file not found: {path}")
        return None

    # Open and resize
    img = Image.open(path)
    img = img.convert("RGB")
    img.thumbnail(max_size)  # keeps aspect ratio, fits into max_size

    # Save to bytes with compression
    buf = BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True)
    img_bytes = buf.getvalue()

    print(f"{path} compressed to ~{len(img_bytes) / 1024:.1f} KB")

    mime_img = MIMEImage(img_bytes, _subtype="jpeg")
    mime_img.add_header("Content-ID", f"<{cid}>")
    return mime_img


# ==========================
# HTML BODY
# ==========================

def build_html_body(recipient_name, tracking_url, pixel_url):
    greeting_name = recipient_name if recipient_name else "there"

    html = f"""
    <html>
    <body style="margin:0; padding:0; background-color:#f7f7f7; font-family:Arial, sans-serif;">
      <table align="center" width="100%" cellpadding="0" cellspacing="0" style="padding:20px 0;">
        <tr>
          <td>
            <table align="center" width="600" cellpadding="0" cellspacing="0"
                   style="background-color:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 2px 10px rgba(0,0,0,0.06);">

              <!-- HEADER -->
              <tr>
                <td style="background:#000000; color:#ffffff; padding:18px 24px; text-align:center;">
                  <h1 style="margin:0; font-size:24px;">La Chic Designs</h1>
                </td>
              </tr>

             <!-- TWO IMAGES SIDE BY SIDE (RESPONSIVE) -->
<tr>
  <td style="padding:0;">

    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
      <tr>

        <!-- LEFT IMAGE -->
        <td align="center" valign="top" style="padding:0; width:50%;">
          <img src="cid:handbag1"
               alt="Handbag 1"
               style="display:block; width:100%; max-width:300px; height:250px; border:0; margin:0 auto;">
        </td>

        <!-- RIGHT IMAGE -->
        <td align="center" valign="top" style="padding:0; width:50%;">
          <img src="cid:handbag2"
               alt="Handbag 2"
               style="display:block; width:100%; max-width:300px; height:250px; border:0; margin:0 auto;">
        </td>

      </tr>
    </table>

  </td>
</tr>


              <!-- MAIN CONTENT -->
              <tr>
                <td style="padding:24px 28px; color:#333;">

                  <p style="margin:0 0 18px; font-size:15px; line-height:1.6;">
                    Hi {greeting_name},
                  </p>

                  <p style="margin:0 0 18px; font-size:15px; line-height:1.6;">
                    I just wanted to follow up on the email I sent last week regarding our
                    <strong> handmade beaded handbags</strong> at La Chic Designs.
                  </p>

                  <p style="margin:0 0 18px; font-size:15px; line-height:1.6;">
                    Many of our retail partners have seen strong sell-through, especially in gift-focused and impulse-buy sections, and
                    I thought our collection could be a great fit for you assortment. Check out our 2025 handbag catalog
                    <a href="{CATALOG_URL}" style="color:#0066cc; text-decoration:underline;">here</a>.
                  </p>

                  <!-- BUTTON ABOVE USERNAME/PASSWORD -->
                  <div style="text-align:center; margin:25px 0;">
                    <a href="{tracking_url}"
                       style="display:inline-block; padding:12px 28px; background:#000; color:#fff;
                              border-radius:999px; text-decoration:none; font-size:14px;">
                      Visit Our Website
                    </a>
                  </div>

                  <!-- LOGIN INFO -->
                  <p style="margin:0 0 18px; font-size:15px; line-height:1.6;">
                    When you access our website, please use:<br>
                    <strong>Username:</strong> User50<br>
                    <strong>Password:</strong> LaChicDesignsUser50!
                  </p>

                   <p style="margin:0 0 18px; font-size:15px; line-height:1.6;">
                    I’d be happy to share pricing, minimums, or answer any questions — just reply
                    to this email and I’ll take care of the rest.
                  </p>

                  <p style="margin:0; font-size:15px; line-height:1.6;">
                    Best regards,<br>
                    <strong>Manik Chhabra</strong><br>
                    Senior Sales Specialist | La Chic Designs<br>
                    <a href="mailto:manik@lachicdesigns.com" style="color:#0066cc; text-decoration:none;">
                      manik@lachicdesigns.com
                    </a> | 469-248-3513 |
                    <a href="https://www.lachicdesigns.com" style="color:#0066cc; text-decoration:none;">
                      www.lachicdesigns.com
                    </a>
                  </p>

                </td>
              </tr>

            </table>

            <!-- TRACKING PIXEL -->
            <img src="{pixel_url}" width="1" height="1" style="display:none;" alt="">
          </td>
        </tr>
      </table>
    </body>
    </html>
    """
    return html


# ==========================
# MESSAGE BUILDER
# ==========================

def create_email_message(to_email, recipient_name):
    tracking_url, pixel_url = build_tracking(to_email)

    msg = MIMEMultipart("related")
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = to_email
    msg["Subject"] = EMAIL_SUBJECT

    alt = MIMEMultipart("alternative")
    msg.attach(alt)

    # Plain-text fallback
    text_body = f"""Hi {recipient_name or ''},

I just wanted to follow up on the email I sent last week regarding our handmade
leather handbags at La Chic Designs.

Many of our retail partners have seen strong sell-through, especially in gift-focused and impulse-buy sections, and
I thought our collection could be a great fit for you assortment. Check out our 2025 handbag catalog.

Catalog:
{CATALOG_URL}

Visit the website:
{tracking_url}

Login:
Username: User50
Password: LaChicDesignsUser50!

I’d be happy to share pricing, minimums, or answer any questions — just reply
to this email and I’ll take care of the rest.

Best regards,
Manik Chhabra
Senior Sales Specialist | La Chic Designs
manik@lachicdesigns.com
469-248-3513
www.lachicdesigns.com
"""
    alt.attach(MIMEText(text_body, "plain"))

    # HTML
    html_body = build_html_body(recipient_name, tracking_url, pixel_url)
    alt.attach(MIMEText(html_body, "html"))

    # INLINE IMAGES (compressed)
    img1 = make_inline_image(HANDBAG_IMAGE_PATH_1, "handbag1")
    if img1:
        msg.attach(img1)

    img2 = make_inline_image(HANDBAG_IMAGE_PATH_2, "handbag2")
    if img2:
        msg.attach(img2)

    return msg


# ==========================
# BULK SEND (BATCHED)
# ==========================

def send_bulk_emails():
    df = pd.read_excel(EXCEL_PATH)

    df = df[df[EMAIL_COLUMN].notna()]
    df[EMAIL_COLUMN] = df[EMAIL_COLUMN].astype(str).str.strip()
    df = df[df[EMAIL_COLUMN] != ""]

    recipients = []
    for _, row in df.iterrows():
        email = row[EMAIL_COLUMN]
        name = row[NAME_COLUMN] if NAME_COLUMN in df.columns and not pd.isna(row[NAME_COLUMN]) else None
        recipients.append((email, name))

    context = ssl.create_default_context()
    batch_size = 15
    total = len(recipients)

    for i in range(0, total, batch_size):
        batch = recipients[i:i+batch_size]
        print(f"\nProcessing batch {i//batch_size + 1} ({len(batch)} emails)")

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls(context=context)
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

            for to_email, name in batch:
                msg = create_email_message(to_email, name)
                try:
                    server.sendmail(EMAIL_ADDRESS, [to_email], msg.as_string())
                    print(f"Sent → {to_email}")
                except Exception as e:
                    print(f"ERROR sending to {to_email}: {e}")

        time.sleep(2)


if __name__ == "__main__":
    send_bulk_emails()









