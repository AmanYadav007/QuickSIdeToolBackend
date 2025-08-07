# Adobe API Quick Setup Guide

Since you already have Client ID, Client Secret, and Organization ID, you just need to complete these final steps:

## **Step 1: Get Your Account ID**

1. Go to [Adobe Developer Console](https://developer.adobe.com/console)
2. Click on your existing project
3. Go to **"Service Account (JWT)"** section
4. Look for **"Account ID"** (this is different from Organization ID)
5. Copy this Account ID

## **Step 2: Generate Private Key**

1. In the same **"Service Account (JWT)"** section
2. Click **"Generate Key Pair"**
3. Download the private key file (`.key` format)
4. Save it in your `backend` folder as `private.key`

## **Step 3: Create Environment File**

Create a file called `.env` in your `backend` directory with this content:

```env
# Adobe API Configuration
ADOBE_CLIENT_ID=your_client_id_here
ADOBE_CLIENT_SECRET=your_client_secret_here
ADOBE_ORGANIZATION_ID=your_organization_id_here
ADOBE_ACCOUNT_ID=your_account_id_here

# Private Key (file path method)
ADOBE_PRIVATE_KEY_PATH=./private.key

# Backend Configuration
REACT_APP_BACKEND_URL=https://quicksidetoolbackend.onrender.com

# Enable Adobe features
ADOBE_ENABLED=true
```

## **Step 4: Replace Placeholder Values**

Replace the placeholder values with your actual credentials:

- `your_client_id_here` → Your actual Client ID
- `your_client_secret_here` → Your actual Client Secret  
- `your_organization_id_here` → Your actual Organization ID
- `your_account_id_here` → Your actual Account ID

## **Step 5: Test the Setup**

1. Make sure your `private.key` file is in the `backend` folder
2. Restart your backend server
3. Test the Adobe endpoints

## **File Structure**
Your backend folder should look like this:
```
backend/
├── .env                    # Your environment variables
├── private.key            # Your private key file
├── app.py                 # Main application
├── adobe_service.py       # Adobe integration
└── requirements.txt       # Dependencies
```

## **Troubleshooting**

If you get authentication errors:
1. Double-check all credentials are correct
2. Make sure the private key file is in the right location
3. Verify the Account ID is from the Service Account (JWT) section
4. Check that all environment variables are set correctly

## **Security Notes**

- Never commit `.env` or `private.key` to version control
- Keep your private key secure
- The `.env` file should be in your `.gitignore`

That's it! You don't need any additional API keys - the credentials you already have are perfect for the PDF Services API.
