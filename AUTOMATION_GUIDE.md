# ATSReady — Deployment & Automation Guide
## "Set and Forget" Wiring Instructions (Stripe Edition)

---

## Architecture Overview

```
User visits site → Uploads PDF → Clicks "Convert Now" button
    ↓
Stripe Checkout opens → User pays £7 (card, Apple Pay, Google Pay)
    ↓
Stripe fires webhook → Make.com scenario triggers
    ↓
Make.com: Retrieves uploaded PDF → Calls Python conversion API
    ↓
Script produces: ATS-optimised .docx + JSON score report
    ↓
Make.com: Emails customer the .docx + score summary
```

---

## Total Cost Breakdown

| Item                     | Cost         | Notes                                    |
|--------------------------|--------------|------------------------------------------|
| Domain (atsready.co.uk)  | ~£8/year     | Porkbun, Namecheap, or Cloudflare        |
| Hosting (static site)    | Free         | Netlify, Cloudflare Pages, or Vercel     |
| Make.com (automation)    | Free tier    | 1,000 operations/month free              |
| Stripe account           | Free         | Takes 1.5% + 20p per UK card transaction |
| Python hosting (script)  | Free–£5/mo   | PythonAnywhere free tier, or Railway.app |
| Email sending            | Free         | Gmail SMTP or Make.com built-in          |
| **TOTAL SETUP**          | **~£8–£15**  | Well under £100 budget                   |

### Stripe vs PayPal Fee Comparison (on a £7 transaction)

| Provider | Fee formula         | You receive |
|----------|---------------------|-------------|
| Stripe   | 1.5% + 20p (UK)     | **£6.70**   |
| PayPal   | 2.9% + 30p          | £6.50       |

Stripe saves you ~20p per transaction and supports Apple Pay / Google Pay
out of the box, which reduces checkout friction significantly.

---

## Step-by-Step Setup

### 1. Domain & Hosting (30 minutes)

1. Register `atsready.co.uk` (or similar) on Porkbun (~£8/year).
2. Create a free account on **Netlify** (netlify.com).
3. Drag-and-drop the `index.html` file to Netlify to deploy.
4. Connect your custom domain in Netlify dashboard → DNS settings.

### 2. Stripe Setup (30 minutes)

Since you already have Stripe experience, this should be quick.

#### 2a. Create a Product and Payment Link

1. Go to **Stripe Dashboard → Products → + Add product**.
2. Set:
   - Name: `ATS Resume Conversion`
   - Price: `£7.00 GBP` (one-time)
3. After creating, go to **Payment Links** (left sidebar).
4. Click **+ New** → select your `ATS Resume Conversion` product.
5. Configure the payment link:
   - **After payment → Confirmation page**: Use a custom URL
     (e.g., `https://atsready.co.uk/thanks.html`)
   - **Collect email address**: Yes (this is how you'll email the result)
   - **Allow promotion codes**: Optional, but useful for marketing
   - **Tax collection**: If you need to charge VAT, enable Stripe Tax
6. **Copy the Payment Link URL** — this goes into your website button.

#### 2b. Set Up Webhooks (for automation)

1. Go to **Stripe Dashboard → Developers → Webhooks**.
2. Click **+ Add endpoint**.
3. Set the endpoint URL to your **Make.com webhook URL** (created in Step 4).
4. Select events to listen to: `checkout.session.completed`
5. Click **Add endpoint**.
6. **Copy the webhook signing secret** — you'll use this in Make.com to
   verify that incoming requests genuinely come from Stripe.

#### 2c. Enable Payment Methods

In **Stripe Dashboard → Settings → Payment methods**, ensure these are on:
- Cards (Visa, Mastercard, Amex)
- Apple Pay (auto-enabled with Stripe)
- Google Pay (auto-enabled with Stripe)
- Link (Stripe's one-click checkout — recommended)

All of these work with no extra setup on your part. Customers on
mobile will automatically see Apple Pay / Google Pay buttons, which
dramatically reduces checkout friction.

#### 2d. Set GBP as Default Currency

Go to **Settings → General → Currency** and confirm GBP is your default.
Any payment links you create will default to GBP. You can always create
additional links in USD, EUR etc. later if you want to target other markets.

### 3. Python Script Hosting (30 minutes)

**Option A: PythonAnywhere (Free tier — recommended to start)**

1. Sign up at pythonanywhere.com (free Beginner account).
2. Upload `ats_converter.py` via the Files tab.
3. Open a Bash console and install dependencies:
   ```bash
   pip install pdfplumber python-docx
   ```
4. Create a simple Flask API wrapper — create `app.py`:

   ```python
   import os
   import json
   import tempfile
   from flask import Flask, request, send_file, jsonify
   from ats_converter import convert_resume

   app = Flask(__name__)

   # Simple API key for security (set this to a random string)
   API_KEY = os.environ.get("ATS_API_KEY", "your-secret-key-change-this")

   @app.route("/convert", methods=["POST"])
   def convert():
       # Verify API key
       if request.headers.get("X-API-Key") != API_KEY:
           return jsonify({"error": "Unauthorized"}), 401

       # Check for file
       if "resume" not in request.files:
           return jsonify({"error": "No file uploaded"}), 400

       file = request.files["resume"]
       if not file.filename.lower().endswith(".pdf"):
           return jsonify({"error": "Only PDF files accepted"}), 400

       # Process in temp directory
       with tempfile.TemporaryDirectory() as tmpdir:
           input_path = os.path.join(tmpdir, "input.pdf")
           output_path = os.path.join(tmpdir, "ats_optimised.docx")

           file.save(input_path)
           result = convert_resume(input_path, output_path)

           if result["success"]:
               return send_file(
                   output_path,
                   as_attachment=True,
                   download_name="ats_optimised_resume.docx",
                   mimetype="application/vnd.openxmlformats-"
                            "officedocument.wordprocessingml.document"
               )
           else:
               return jsonify(result), 422

   @app.route("/health", methods=["GET"])
   def health():
       return jsonify({"status": "ok"})
   ```

5. In PythonAnywhere → Web tab → create a new web app using Flask.
6. Point it to your `app.py`.
7. Set the environment variable `ATS_API_KEY` in the WSGI config file.

**Option B: Railway.app (more scalable, free tier available)**

1. Push both files to a GitHub repo.
2. Connect Railway to GitHub → auto-deploys.
3. Add a `requirements.txt`:
   ```
   flask
   pdfplumber
   python-docx
   gunicorn
   ```
4. Add a `Procfile`:
   ```
   web: gunicorn app:app
   ```

### 4. Make.com Automation (45 minutes)

This is where the "set and forget" magic happens.

**Create a Make.com account** (free, 1000 ops/month).

**Scenario: Stripe Payment → Convert → Email Result**

#### Module 1: Stripe Webhook (Trigger)

1. Create a new Scenario in Make.com.
2. Add **Stripe → Watch Events** as the trigger module.
3. Connect your Stripe account (Make.com will ask for your API key —
   use the **Restricted key** from Stripe Dashboard → Developers → API keys).
4. Set the event type to: `checkout.session.completed`
5. This fires every time a customer successfully pays.

The webhook payload includes:
- `customer_details.email` — the customer's email address
- `payment_intent` — payment reference
- `metadata` — any custom data you attached (e.g., file reference)

#### Module 2: Retrieve the Uploaded PDF

For the MVP, the simplest approach is this two-part flow:

**How the file gets to you:**

On the website, after the user selects their file and clicks "Convert Now",
the flow is:

1. File is uploaded to **temporary storage** BEFORE redirecting to Stripe.
2. A unique file ID is stored and passed to Stripe as `metadata`.
3. When the Stripe webhook fires, Make.com uses the file ID to retrieve it.

**Temporary storage options (all free):**

- **file.io** — Files auto-delete after first download. Free API.
  Upload via `POST https://file.io` with the file as form data.
  Returns a download URL.

- **Uploadcare** — Free tier: 3,000 uploads/month.
  Has a JavaScript widget you can embed directly in your page.
  Returns a CDN URL for the file.

- **Your own PythonAnywhere endpoint** — Add a `/upload` route to your
  Flask app that stores files temporarily and returns an ID.

**Recommended: Uploadcare** (simplest integration with the frontend).

In Make.com:
1. Add an **HTTP → Get a File** module.
2. URL: The file URL from the Stripe webhook's `metadata.file_url` field.
3. This downloads the PDF into the Make.com pipeline.

#### Module 3: Call the Conversion API

1. Add an **HTTP → Make a Request** module.
2. Configure:
   - URL: `https://yourusername.pythonanywhere.com/convert`
   - Method: `POST`
   - Headers: `X-API-Key: your-secret-key`
   - Body type: `multipart/form-data`
   - Field name: `resume`
   - Field value: Map the file from Module 2
3. Parse response as: Binary data (the .docx file)

#### Module 4: Email the Result

1. Add a **Gmail → Send an Email** module (or any SMTP module).
2. Configure:
   - **To**: `{{customer_details.email}}` from the Stripe webhook
   - **Subject**: `Your ATS-Optimised Resume is Ready`
   - **Body**:
     ```
     Hi there,

     Great news — your resume has been converted to an ATS-friendly format.

     Please find your optimised resume attached as a .docx file.
     This format is designed to pass through Applicant Tracking Systems
     used by most major employers.

     Tips for best results:
     - Upload this .docx file directly to job applications
     - Avoid converting it back to PDF for online submissions
     - Keep the clean formatting — don't add tables or graphics

     Good luck with your job search!

     — The ATSReady Team
     ```
   - **Attachment**: The .docx file from Module 3

#### Module 5: Log the Transaction

1. Add a **Google Sheets → Add Row** module.
2. Map: Date, customer email, Stripe payment ID, conversion success/fail.
3. This gives you a simple dashboard of all orders.

#### Error Handling

1. Add an **Error Handler** route on the HTTP module (Module 3).
2. If conversion fails:
   - Send a polite error email to the customer explaining the issue
   - Log the failure in your Google Sheet
   - Flag for manual review
3. Consider adding a **Router** after Module 1 to handle refund scenarios.

---

## Complete User Journey

```
1. User lands on atsready.co.uk
2. Scrolls down, uploads their PDF resume
3. File is uploaded to temporary storage (Uploadcare or file.io)
4. User clicks "Convert Now — £7" button
5. Stripe Checkout opens (supports card, Apple Pay, Google Pay)
6. User pays → redirected to thank-you page
7. Stripe webhook fires → Make.com scenario triggers
8. Make.com downloads PDF → calls conversion API → gets .docx
9. Make.com emails customer the .docx + ATS score
10. Customer receives result in 2-5 minutes
```

---

## Alternative: Even Simpler MVP (Validate Demand First)

Before wiring up full automation, you can validate demand manually:

1. Deploy the landing page with a Stripe Payment Link button.
2. Ask customers to email their PDF to convert@atsready.co.uk after paying.
3. You receive the Stripe payment notification on your phone.
4. Run the Python script manually on PythonAnywhere.
5. Email the result back.

This lets you confirm people will actually pay before spending time on
automation. Once you're getting 2-3 orders per day, wire up Make.com.

---

## Stripe-Specific Tips

### Testing Before Launch

1. Use **Stripe Test Mode** (toggle in top-right of dashboard).
2. Test card number: `4242 4242 4242 4242` (any future expiry, any CVC).
3. Create a test Payment Link and run through the full flow.
4. Stripe test webhooks can be triggered manually from the dashboard:
   Developers → Webhooks → Select endpoint → Send test webhook.

### Stripe Payment Link Customisation

- Add your logo and brand colour in **Settings → Branding**.
- Stripe Checkout automatically adapts to mobile screens.
- Apple Pay and Google Pay buttons appear automatically on supported devices.
- You can add a custom "thank you" redirect URL.

### Monitoring

- **Stripe Dashboard → Payments**: See all transactions in real-time.
- Set up **Stripe email notifications** for successful payments.
- The Stripe mobile app gives you push notifications for each sale.

---

## Scaling Notes

- **Free tier limits**: PythonAnywhere free = 1 web app, limited CPU.
  Fine for ~10-20 conversions/day. Upgrade to £5/mo Hacker plan for more.
- **Make.com free tier**: 1,000 operations/month ≈ ~200 conversions/month
  (each conversion uses ~5 operations). Upgrade at £9/mo for 10,000 ops.
- **Stripe**: No monthly fees. Just 1.5% + 20p per UK transaction.
- **Revenue at 10 conversions/day**: 10 × £7 = £70/day = ~£2,100/month.
  Stripe fees ~£40/mo. Hosting ~£5-15/mo. **Net profit ~£2,000/month**.

---

## Marketing (Zero-Cost Channels)

1. **Reddit**: Post helpful resume advice in r/jobs, r/UKjobs, r/resumes
   with a subtle link. DON'T spam — provide genuine value first.
2. **TikTok/Reels**: "Did you know 75% of resumes are rejected by robots?"
   Short-form content performs extremely well in the job-seeker niche.
3. **SEO**: Blog posts on the site targeting long-tail keywords:
   "What is an ATS?", "Why your PDF resume gets rejected",
   "ATS-friendly resume format guide UK".
4. **Fiverr/Upwork**: List this as a service at £10-15, run the tool,
   deliver in minutes. Use it as a lead-gen funnel back to the site.
5. **LinkedIn**: Post about ATS tips. The job-seeker audience is literally
   on the platform already.
