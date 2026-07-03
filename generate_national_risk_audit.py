import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
businesses = [
    {
        "id": "BIZ-01",
        "name": "Leopold Cafe & Bar",
        "category": "Restaurants & Cafes",
        "location": "Mumbai",
        "rating": 4.2,
        "reviews": 8400,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "http://leopoldcafe.com",
        "https_redirect": "No",
        "ssl_status": "None",
        "ssl_provider": "None",
        "admin_panel_exposed": "Yes",
        "forms_protection": "Unprotected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "X-Frame-Options, CSP, HSTS",
    },
    {
        "id": "BIZ-02",
        "name": "Toit Brewpub",
        "category": "Restaurants & Cafes",
        "location": "Bengaluru",
        "rating": 4.6,
        "reviews": 9800,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://toit.in",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "Let's Encrypt",
        "admin_panel_exposed": "Yes",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "CSP, HSTS",
    },
    {
        "id": "BIZ-03",
        "name": "Glen's Bakehouse",
        "category": "Restaurants & Cafes",
        "location": "Bengaluru",
        "rating": 4.4,
        "reviews": 3200,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://glensbakehouse.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "Cloudflare",
        "admin_panel_exposed": "No",
        "forms_protection": "Unprotected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-04",
        "name": "Flurys",
        "category": "Restaurants & Cafes",
        "location": "Kolkata",
        "rating": 4.3,
        "reviews": 4500,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "http://flurys.com",
        "https_redirect": "No",
        "ssl_status": "None",
        "ssl_provider": "None",
        "admin_panel_exposed": "No",
        "forms_protection": "Unprotected",
        "directory_listing": "Yes",
        "safe_browsing": "Clean",
        "missing_headers": "All Headers Missing",
    },
    {
        "id": "BIZ-05",
        "name": "Mavalli Tiffin Room (MTR)",
        "category": "Restaurants & Cafes",
        "location": "Bengaluru",
        "rating": 4.5,
        "reviews": 12000,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://mtrrestaurants.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-06",
        "name": "Paradise Biryani",
        "category": "Restaurants & Cafes",
        "location": "Hyderabad",
        "rating": 4.0,
        "reviews": 18500,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://paradisebiryani.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "Sectigo",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "X-Frame-Options",
    },
    {
        "id": "BIZ-07",
        "name": "Peter Cat",
        "category": "Restaurants & Cafes",
        "location": "Kolkata",
        "rating": 4.5,
        "reviews": 8900,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "http://petercatkolkata.co.in",
        "https_redirect": "No",
        "ssl_status": "Expired",
        "ssl_provider": "Let's Encrypt (Expired)",
        "admin_panel_exposed": "Yes",
        "forms_protection": "Unprotected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "All Headers Missing",
    },
    {
        "id": "BIZ-08",
        "name": "Cafe Goodluck",
        "category": "Restaurants & Cafes",
        "location": "Pune",
        "rating": 4.3,
        "reviews": 6200,
        "gbp_claimed": "No",
        "website_status": "No Website",
        "website_url": "None",
        "https_redirect": "N/A",
        "ssl_status": "None",
        "ssl_provider": "None",
        "admin_panel_exposed": "N/A",
        "forms_protection": "N/A",
        "directory_listing": "N/A",
        "safe_browsing": "N/A",
        "missing_headers": "N/A",
    },
    {
        "id": "BIZ-09",
        "name": "Britto's Restaurant & Bar",
        "category": "Restaurants & Cafes",
        "location": "Goa",
        "rating": 4.1,
        "reviews": 11500,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "http://brittosgoa.com",
        "https_redirect": "No",
        "ssl_status": "None",
        "ssl_provider": "None",
        "admin_panel_exposed": "No",
        "forms_protection": "Unprotected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "X-Frame-Options, CSP",
    },
    {
        "id": "BIZ-10",
        "name": "Karim's Restaurant",
        "category": "Restaurants & Cafes",
        "location": "Delhi",
        "rating": 4.2,
        "reviews": 5600,
        "gbp_claimed": "No",
        "website_status": "Active",
        "website_url": "http://karimsdelhi.in",
        "https_redirect": "No",
        "ssl_status": "None",
        "ssl_provider": "None",
        "admin_panel_exposed": "Yes",
        "forms_protection": "Unprotected",
        "directory_listing": "Yes",
        "safe_browsing": "Clean",
        "missing_headers": "All Headers Missing",
    },
    {
        "id": "BIZ-11",
        "name": "Apollo Hospitals Greams Road",
        "category": "Healthcare & Clinics",
        "location": "Chennai",
        "rating": 4.2,
        "reviews": 4200,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://apollohospitals.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-12",
        "name": "Fortis Hospital Mulund",
        "category": "Healthcare & Clinics",
        "location": "Mumbai",
        "rating": 4.1,
        "reviews": 2800,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://fortishealthcare.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "Entrust",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-13",
        "name": "Manipal Hospital Old Airport Rd",
        "category": "Healthcare & Clinics",
        "location": "Bengaluru",
        "rating": 4.3,
        "reviews": 3500,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://manipalhospitals.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "Sectigo",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-14",
        "name": "Lilavati Hospital & Research Centre",
        "category": "Healthcare & Clinics",
        "location": "Mumbai",
        "rating": 4.0,
        "reviews": 1900,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://lilavatihospital.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "GlobalSign",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "CSP",
    },
    {
        "id": "BIZ-15",
        "name": "Narayana Health City",
        "category": "Healthcare & Clinics",
        "location": "Bengaluru",
        "rating": 4.4,
        "reviews": 3100,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://narayanahealth.org",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-16",
        "name": "Ganga Hospital",
        "category": "Healthcare & Clinics",
        "location": "Coimbatore",
        "rating": 4.6,
        "reviews": 1500,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "http://gangahospital.co.in",
        "https_redirect": "No",
        "ssl_status": "None",
        "ssl_provider": "None",
        "admin_panel_exposed": "Yes",
        "forms_protection": "Unprotected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "All Headers Missing",
    },
    {
        "id": "BIZ-17",
        "name": "Medanta The Medicity",
        "category": "Healthcare & Clinics",
        "location": "Delhi",
        "rating": 4.3,
        "reviews": 8900,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://medanta.org",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-18",
        "name": "KIMS Hospitals Secunderabad",
        "category": "Healthcare & Clinics",
        "location": "Hyderabad",
        "rating": 4.1,
        "reviews": 2400,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://kimshospitals.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "Sectigo",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-19",
        "name": "Dr. Mehta's Hospitals",
        "category": "Healthcare & Clinics",
        "location": "Chennai",
        "rating": 3.9,
        "reviews": 950,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "http://mehtahospitals.com",
        "https_redirect": "No",
        "ssl_status": "Expired",
        "ssl_provider": "Let's Encrypt (Expired)",
        "admin_panel_exposed": "No",
        "forms_protection": "Unprotected",
        "directory_listing": "No",
        "safe_browsing": "Suspicious",
        "missing_headers": "All Headers Missing",
    },
    {
        "id": "BIZ-20",
        "name": "City Dental Clinic & Implant Center",
        "category": "Healthcare & Clinics",
        "location": "Ahmedabad",
        "rating": 4.2,
        "reviews": 120,
        "gbp_claimed": "No",
        "website_status": "No Website",
        "website_url": "None",
        "https_redirect": "N/A",
        "ssl_status": "None",
        "ssl_provider": "None",
        "admin_panel_exposed": "N/A",
        "forms_protection": "N/A",
        "directory_listing": "N/A",
        "safe_browsing": "N/A",
        "missing_headers": "N/A",
    },
    {
        "id": "BIZ-21",
        "name": "Nalli Silks T-Nagar",
        "category": "Retail & Shops",
        "location": "Chennai",
        "rating": 4.5,
        "reviews": 3200,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://nalli.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "GeoTrust",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "CSP",
    },
    {
        "id": "BIZ-22",
        "name": "Bahrisons Booksellers Khan Mkt",
        "category": "Retail & Shops",
        "location": "Delhi",
        "rating": 4.7,
        "reviews": 890,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://bahrisons.in",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "Sectigo",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "CSP, HSTS",
    },
    {
        "id": "BIZ-23",
        "name": "Sabyasachi Flagship Store Kolkata",
        "category": "Retail & Shops",
        "location": "Kolkata",
        "rating": 4.6,
        "reviews": 650,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://sabyasachi.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-24",
        "name": "Sarojini Saree House",
        "category": "Retail & Shops",
        "location": "Jaipur",
        "rating": 3.6,
        "reviews": 45,
        "gbp_claimed": "No",
        "website_status": "No Website",
        "website_url": "None",
        "https_redirect": "N/A",
        "ssl_status": "None",
        "ssl_provider": "None",
        "admin_panel_exposed": "N/A",
        "forms_protection": "N/A",
        "directory_listing": "N/A",
        "safe_browsing": "N/A",
        "missing_headers": "N/A",
    },
    {
        "id": "BIZ-25",
        "name": "Mysore Silk Udyog",
        "category": "Retail & Shops",
        "location": "Bengaluru",
        "rating": 4.1,
        "reviews": 1200,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "http://mysoresilkudyog.com",
        "https_redirect": "No",
        "ssl_status": "None",
        "ssl_provider": "None",
        "admin_panel_exposed": "No",
        "forms_protection": "Unprotected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "X-Frame-Options",
    },
    {
        "id": "BIZ-26",
        "name": "Fabindia Fort",
        "category": "Retail & Shops",
        "location": "Mumbai",
        "rating": 4.2,
        "reviews": 980,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://fabindia.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-27",
        "name": "Manyavar",
        "category": "Retail & Shops",
        "location": "Kolkata",
        "rating": 4.4,
        "reviews": 1500,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://manyavar.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "Sectigo",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-28",
        "name": "Kalaniketan Colaba",
        "category": "Retail & Shops",
        "location": "Mumbai",
        "rating": 4.0,
        "reviews": 720,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "http://kalaniketancolaba.in",
        "https_redirect": "No",
        "ssl_status": "None",
        "ssl_provider": "None",
        "admin_panel_exposed": "Yes",
        "forms_protection": "Unprotected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "All Headers Missing",
    },
    {
        "id": "BIZ-29",
        "name": "Ratnadeep Supermarket Jubilee Hills",
        "category": "Retail & Shops",
        "location": "Hyderabad",
        "rating": 4.3,
        "reviews": 1100,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://ratnadeep.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "Cloudflare",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "CSP",
    },
    {
        "id": "BIZ-30",
        "name": "Coimbatore Grocery Mart",
        "category": "Retail & Shops",
        "location": "Coimbatore",
        "rating": 3.7,
        "reviews": 35,
        "gbp_claimed": "No",
        "website_status": "No Website",
        "website_url": "None",
        "https_redirect": "N/A",
        "ssl_status": "None",
        "ssl_provider": "None",
        "admin_panel_exposed": "N/A",
        "forms_protection": "N/A",
        "directory_listing": "N/A",
        "safe_browsing": "N/A",
        "missing_headers": "N/A",
    },
    {
        "id": "BIZ-31",
        "name": "Cyril Amarchand Mangaldas",
        "category": "Professional Services",
        "location": "Mumbai",
        "rating": 4.5,
        "reviews": 210,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://cyrilshroff.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "CSP, HSTS",
    },
    {
        "id": "BIZ-32",
        "name": "AZB & Partners Corporate Office",
        "category": "Professional Services",
        "location": "Mumbai",
        "rating": 4.6,
        "reviews": 180,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://azbpartners.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-33",
        "name": "Trilegal Advocates",
        "category": "Professional Services",
        "location": "Delhi",
        "rating": 4.4,
        "reviews": 150,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://trilegal.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "Sectigo",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "CSP",
    },
    {
        "id": "BIZ-34",
        "name": "Khaitan & Co Advocates",
        "category": "Professional Services",
        "location": "Kolkata",
        "rating": 4.5,
        "reviews": 190,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://khaitanco.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-35",
        "name": "Luthra & Luthra Law Offices Delhi",
        "category": "Professional Services",
        "location": "Delhi",
        "rating": 4.3,
        "reviews": 140,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://luthra.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-36",
        "name": "J. Sagar Associates",
        "category": "Professional Services",
        "location": "Delhi",
        "rating": 4.2,
        "reviews": 120,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://jsalaw.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "Sectigo",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-37",
        "name": "Singhania & Co. Legal Advisors",
        "category": "Professional Services",
        "location": "Bengaluru",
        "rating": 4.1,
        "reviews": 85,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://singhania.in",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "Sectigo",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "CSP",
    },
    {
        "id": "BIZ-38",
        "name": "Gupta Chartered Accountants",
        "category": "Professional Services",
        "location": "Pune",
        "rating": 3.8,
        "reviews": 12,
        "gbp_claimed": "No",
        "website_status": "No Website",
        "website_url": "None",
        "https_redirect": "N/A",
        "ssl_status": "None",
        "ssl_provider": "None",
        "admin_panel_exposed": "N/A",
        "forms_protection": "N/A",
        "directory_listing": "N/A",
        "safe_browsing": "N/A",
        "missing_headers": "N/A",
    },
    {
        "id": "BIZ-39",
        "name": "Mehta Tax Advisors",
        "category": "Professional Services",
        "location": "Mumbai",
        "rating": 3.7,
        "reviews": 18,
        "gbp_claimed": "No",
        "website_status": "No Website",
        "website_url": "None",
        "https_redirect": "N/A",
        "ssl_status": "None",
        "ssl_provider": "None",
        "admin_panel_exposed": "N/A",
        "forms_protection": "N/A",
        "directory_listing": "N/A",
        "safe_browsing": "N/A",
        "missing_headers": "N/A",
    },
    {
        "id": "BIZ-40",
        "name": "Sharma Real Estate Consultants",
        "category": "Professional Services",
        "location": "Jaipur",
        "rating": 3.5,
        "reviews": 8,
        "gbp_claimed": "No",
        "website_status": "No Website",
        "website_url": "None",
        "https_redirect": "N/A",
        "ssl_status": "None",
        "ssl_provider": "None",
        "admin_panel_exposed": "N/A",
        "forms_protection": "N/A",
        "directory_listing": "N/A",
        "safe_browsing": "N/A",
        "missing_headers": "N/A",
    },
    {
        "id": "BIZ-41",
        "name": "Taj Mahal Palace Colaba",
        "category": "Hospitality & Hotels",
        "location": "Mumbai",
        "rating": 4.9,
        "reviews": 18500,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://tajhotels.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-42",
        "name": "The Oberoi Mahatma Gandhi Rd",
        "category": "Hospitality & Hotels",
        "location": "Bengaluru",
        "rating": 4.8,
        "reviews": 6200,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://oberoihotels.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-43",
        "name": "The Leela Palace MRC Nagar",
        "category": "Hospitality & Hotels",
        "location": "Chennai",
        "rating": 4.7,
        "reviews": 4800,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://theleela.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-44",
        "name": "ITC Grand Chola Guindy",
        "category": "Hospitality & Hotels",
        "location": "Chennai",
        "rating": 4.8,
        "reviews": 9500,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://itchotels.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "Sectigo",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-45",
        "name": "JW Marriott Senapati Bapat Rd",
        "category": "Hospitality & Hotels",
        "location": "Pune",
        "rating": 4.7,
        "reviews": 5400,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://marriott.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-46",
        "name": "Rambagh Palace",
        "category": "Hospitality & Hotels",
        "location": "Jaipur",
        "rating": 4.9,
        "reviews": 3200,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://tajhotels.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-47",
        "name": "Wildflower Hall Anamika",
        "category": "Hospitality & Hotels",
        "location": "Shimla",
        "rating": 4.8,
        "reviews": 1100,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://oberoihotels.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "DigiCert",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-48",
        "name": "Brunton Boatyard",
        "category": "Hospitality & Hotels",
        "location": "Kochi",
        "rating": 4.5,
        "reviews": 450,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "https://cghearth.com",
        "https_redirect": "Yes",
        "ssl_status": "Valid",
        "ssl_provider": "GeoTrust",
        "admin_panel_exposed": "No",
        "forms_protection": "Protected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "None",
    },
    {
        "id": "BIZ-49",
        "name": "Gokulam Park Plaza",
        "category": "Hospitality & Hotels",
        "location": "Kochi",
        "rating": 4.0,
        "reviews": 750,
        "gbp_claimed": "Yes",
        "website_status": "Active",
        "website_url": "http://gokulampark.co.in",
        "https_redirect": "No",
        "ssl_status": "Expired",
        "ssl_provider": "Let's Encrypt (Expired)",
        "admin_panel_exposed": "No",
        "forms_protection": "Unprotected",
        "directory_listing": "No",
        "safe_browsing": "Clean",
        "missing_headers": "All Headers Missing",
    },
    {
        "id": "BIZ-50",
        "name": "Shimla Residency Guest House",
        "category": "Hospitality & Hotels",
        "location": "Shimla",
        "rating": 3.5,
        "reviews": 45,
        "gbp_claimed": "No",
        "website_status": "No Website",
        "website_url": "None",
        "https_redirect": "N/A",
        "ssl_status": "None",
        "ssl_provider": "None",
        "admin_panel_exposed": "N/A",
        "forms_protection": "N/A",
        "directory_listing": "N/A",
        "safe_browsing": "N/A",
        "missing_headers": "N/A",
    }
]
os.makedirs("charts", exist_ok=True)
processed_data = []
for biz in businesses:
    row = biz.copy()
    risk_indicators_list = []
    if row["website_status"] == "No Website":
        if row["gbp_claimed"] == "No":
            row["security_score"] = 0
            row["key_security_gaps"] = "Digital Absence, Unclaimed Google Business Profile"
        else:
            row["security_score"] = 1
            row["key_security_gaps"] = "Digital Absence, No Website Present"
        risk_indicators_list.append("No Website (High Risk)")
    else:
        score = 5
        gaps = []
        if row["https_redirect"] == "No":
            score -= 1
            gaps.append("No HTTPS Redirect")
            risk_indicators_list.append("No HTTPS / Not Secure")
        if row["ssl_status"] in ["Expired", "None"]:
            score -= 1
            gaps.append(f"SSL Status: {row['ssl_status']}")
            risk_indicators_list.append("SSL Certificate Expired/Invalid" if row["ssl_status"] == "Expired" else "No SSL Certificate")
        if row["admin_panel_exposed"] == "Yes":
            score -= 1
            gaps.append("Exposed Admin Panel (/wp-admin or /admin)")
            risk_indicators_list.append("Admin Panel Exposed")
        if row["forms_protection"] == "Unprotected" or row["directory_listing"] == "Yes":
            score -= 1
            subgaps = []
            if row["forms_protection"] == "Unprotected":
                subgaps.append("Unprotected Forms")
                risk_indicators_list.append("Forms Without Protection")
            if row["directory_listing"] == "Yes":
                subgaps.append("Directory Listing Enabled")
                risk_indicators_list.append("Open Directory Exposed")
            gaps.append(", ".join(subgaps))
        if row["safe_browsing"] in ["Suspicious", "Blacklisted"]:
            score -= 1
            gaps.append(f"Safe Browsing: {row['safe_browsing']}")
            risk_indicators_list.append("Malware/Phishing Risk Detected")
        if score < 1:
            score = 1
        row["security_score"] = score
        row["key_security_gaps"] = ", ".join(gaps) if gaps else "None (All Secure)"
    if row["security_score"] == 5:
        row["risk_level"] = "Very Secure"
    elif row["security_score"] == 4:
        row["risk_level"] = "Secure"
    elif row["security_score"] == 3:
        row["risk_level"] = "Moderate Risk"
    elif row["security_score"] == 2:
        row["risk_level"] = "High Risk"
    else:
        row["risk_level"] = "Very High Risk"
    row["risk_indicators"] = risk_indicators_list
    rec = ""
    if row["website_status"] == "No Website":
        if row["gbp_claimed"] == "No":
            rec = "1. Claim Google Business Profile immediately to prevent hijacking. 2. Build a secure, mobile-friendly landing page with HTTPS. 3. Standardize NAP (Name, Address, Phone) details."
        else:
            rec = "1. Invest in a basic secure website (HTTPS) to capture organic search traffic. 2. Provide customer inquiries via webforms instead of just phone. 3. Setup official email domain."
    else:
        recs = []
        if row["https_redirect"] == "No" or row["ssl_status"] in ["Expired", "None"]:
            recs.append("Install valid SSL certificate (e.g. Let's Encrypt) and force HTTPS redirect")
        if row["admin_panel_exposed"] == "Yes":
            recs.append("Restrict access to admin pages (/wp-admin, /admin) by IP address, change default URL path, and enable MFA")
        if row["directory_listing"] == "Yes":
            recs.append("Disable directory indexing in server config (.htaccess or Nginx config)")
        if row["forms_protection"] == "Unprotected":
            recs.append("Implement CSRF tokens, secure input sanitization, and captchas on contact forms")
        if row["safe_browsing"] in ["Suspicious", "Blacklisted"]:
            recs.append("Run security malware scans, remove malicious scripts, and request Google review of blacklist")
        if not recs:
            recs.append("Maintain existing secure posture. Schedule routine security auditing and patch CMS regularly.")
        rec = "; ".join(recs)
    row["recommendation"] = rec
    if row["security_score"] <= 2:
        row["outreach_priority"] = "CRITICAL"
    elif row["security_score"] == 3:
        row["outreach_priority"] = "MEDIUM"
    else:
        row["outreach_priority"] = "LOW"
    if row["website_status"] == "No Website":
        if row["gbp_claimed"] == "No":
            row["sales_pitch_hook"] = "Profile Hijacking Risk (Unclaimed Google Maps Profile) & Digital Absence"
        else:
            row["sales_pitch_hook"] = "Digital Absence (Missing Website Customer Acquisition Channel)"
    else:
        gaps_pitch = []
        if row["ssl_status"] in ["Expired", "None"] or row["https_redirect"] == "No":
            gaps_pitch.append("Insecure Protocol / Missing SSL (Active Browser Security Warnings)")
        if row["admin_panel_exposed"] == "Yes":
            gaps_pitch.append("Exposed WordPress Admin Panel (Brute Force Vulnerability)")
        if row["gbp_claimed"] == "No":
            gaps_pitch.append("Unclaimed Business Profile on Google Maps (Malicious Takeover Risk)")
        if row["directory_listing"] == "Yes":
            gaps_pitch.append("Exposed Web Directory (Insecure Server File Listings)")
        if row["forms_protection"] == "Unprotected":
            gaps_pitch.append("Insecure Webforms (Vulnerable to SQLi and Contact Spam)")
        if row["safe_browsing"] in ["Suspicious", "Blacklisted"]:
            gaps_pitch.append("Phishing/Malware Flags Detected on Domain")
        if gaps_pitch:
            row["sales_pitch_hook"] = " | ".join(gaps_pitch)
        else:
            row["sales_pitch_hook"] = "General Maintenance & Continuous Security Penetration Scanning"
    if row["security_score"] <= 2:
        if row["reviews"] >= 1000:
            row["estimated_lead_value"] = "$3,500 - $5,000 (Premium Retainer)"
        elif row["reviews"] >= 100:
            row["estimated_lead_value"] = "$1,500 - $2,500 (Standard Setup)"
        else:
            row["estimated_lead_value"] = "$750 - $1,200 (Essential Setup)"
    elif row["security_score"] == 3:
        if row["reviews"] >= 100:
            row["estimated_lead_value"] = "$1,000 - $1,800 (Remediation Package)"
        else:
            row["estimated_lead_value"] = "$500 - $1,000 (Patch Package)"
    else:
        row["estimated_lead_value"] = "$250 (Routine Audit / Audit Retainer)"
    if row["website_status"] == "No Website":
        row["outreach_strategy"] = "Local Business Visit (Claim Google Maps Listing Pitch)"
    elif row["reviews"] >= 1000 and row["security_score"] <= 2:
        row["outreach_strategy"] = "In-Person Executive Meeting (Deliver Printed Security Briefing)"
    elif row["security_score"] <= 2:
        row["outreach_strategy"] = "Direct Email Outreach + Safe Browsing Alert PDF"
    else:
        row["outreach_strategy"] = "LinkedIn Outreach to Owner / Operations Manager"
    name_clean = row["name"]
    if row["ssl_status"] == "Expired":
        row["outreach_subject_line"] = f"Urgent: Expired SSL Certificate security warning for {name_clean}"
    elif row["ssl_status"] == "None" or row["https_redirect"] == "No":
        if row["website_status"] != "No Website":
            row["outreach_subject_line"] = f"Critical Security Warning: Insecure connection on {name_clean}"
        else:
            row["outreach_subject_line"] = f"Unsecured local presence for {name_clean}"
    elif row["admin_panel_exposed"] == "Yes":
        row["outreach_subject_line"] = f"Security Risk Alert: Exposed WordPress admin panel on {name_clean}"
    elif row["gbp_claimed"] == "No":
        row["outreach_subject_line"] = f"Action Required: Unclaimed Google Maps Profile for {name_clean}"
    else:
        row["outreach_subject_line"] = f"Digital Security Posture Review for {name_clean}"
    if row["website_status"] == "No Website":
        row["first_touch_pitch"] = f"Your Google Maps profile for {name_clean} is currently unclaimed, meaning anyone can change your phone number or edit your address. We can secure it today."
    elif row["ssl_status"] == "Expired":
        row["first_touch_pitch"] = f"Your website ({row['website_url']}) currently triggers a red 'Your Connection is Not Private' warning in browsers. We can restore a secure SSL certificate within 2 hours."
    elif row["ssl_status"] == "None" or row["https_redirect"] == "No":
        row["first_touch_pitch"] = f"Your website ({row['website_url']}) is sending customer passwords and inputs in cleartext without forcing HTTPS, prompting security warnings. We can fix it today."
    elif row["admin_panel_exposed"] == "Yes":
        row["first_touch_pitch"] = f"Your administrative login page is accessible to anyone at {row['website_url']}/wp-admin, leaving you open to brute-force credential stuffing. We can restrict this immediately."
    elif row["gbp_claimed"] == "No":
        row["first_touch_pitch"] = f"While your website is secure, your Google Maps listing is currently unclaimed, exposing you to profile hijacking. We can lock down your Maps listing."
    else:
        row["first_touch_pitch"] = f"We have run a digital presence scan on {name_clean} and verified that your primary protocols are secure. We offer routine scanning packages to maintain this posture."
    processed_data.append(row)
processed_data.sort(key=lambda x: (x["security_score"], -x["reviews"]))
df = pd.DataFrame(processed_data)
csv_columns_order = [
    "id", "name", "category", "location", "rating", "reviews", "gbp_claimed", 
    "website_status", "website_url", "https_redirect", "ssl_status", "ssl_provider", 
    "admin_panel_exposed", "forms_protection", "directory_listing", "safe_browsing", 
    "missing_headers", "security_score", "risk_level", "outreach_priority", 
    "estimated_lead_value", "outreach_strategy", "outreach_subject_line", 
    "first_touch_pitch", "sales_pitch_hook", "key_security_gaps", "recommendation"
]
import csv
df_csv = df[csv_columns_order].copy()
df_csv.columns = [
    col.replace("_", " ").title()
    .replace("Id", "ID")
    .replace("Gbp", "GBP")
    .replace("Https", "HTTPS")
    .replace("Ssl", "SSL")
    for col in df_csv.columns
]
df_csv.to_csv("National_Digital_Risk_Audit_Database.csv", index=False, quoting=csv.QUOTE_ALL)
print("Processed security data saved to National_Digital_Risk_Audit_Database.csv")
print("Configuring dark theme for charts...")
def setup_dark_matplotlib_theme():
    plt.rcParams['figure.facecolor'] = '#120F1D'
    plt.rcParams['axes.facecolor'] = '#120F1D'
    plt.rcParams['text.color'] = '#FAF6F0'
    plt.rcParams['axes.labelcolor'] = '#FAF6F0'
    plt.rcParams['xtick.color'] = '#D5C3C6'
    plt.rcParams['ytick.color'] = '#D5C3C6'
    plt.rcParams['grid.color'] = '#312C45' 
    plt.rcParams['grid.alpha'] = 0.5
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.sans-serif'] = ['Segoe UI', 'Arial', 'DejaVu Sans']
    plt.rcParams['font.size'] = 12
    plt.rcParams['axes.titlesize'] = 16
    plt.rcParams['axes.labelsize'] = 14
    plt.rcParams['xtick.labelsize'] = 12
    plt.rcParams['ytick.labelsize'] = 12
    plt.rcParams['legend.fontsize'] = 11
setup_dark_matplotlib_theme()
cyber_colors = {
    "Very Secure": "#88D49E",    
    "Secure": "#2E8B57",         
    "Moderate Risk": "#F1C40F",   
    "High Risk": "#E67E22",      
    "Very High Risk": "#F07167"  
}
print("Generating distribution pie chart...")
risk_counts = df["risk_level"].value_counts().reindex(["Very Secure", "Secure", "Moderate Risk", "High Risk", "Very High Risk"]).fillna(0)
risk_counts = risk_counts[risk_counts > 0] 
fig, ax = plt.subplots(figsize=(7, 6))
colors = [cyber_colors[label] for label in risk_counts.index]
wedges, texts, autotexts = ax.pie(
    risk_counts, 
    autopct='%1.0f%%', 
    startangle=140, 
    colors=colors,
    wedgeprops=dict(width=0.45, edgecolor='#120F1D', linewidth=4),
    pctdistance=0.75  
)
for autotext in autotexts:
    autotext.set_color('#111827')
    autotext.set_weight('bold')
    autotext.set_fontsize(12)
ax.text(0, 0, f'{len(df)}\nTargets', ha='center', va='center', fontsize=20, fontweight='bold', color='#E5A88B')
ax.legend(wedges, risk_counts.index, title="Risk Level", loc="center left", bbox_to_anchor=(0.95, 0.5), 
          facecolor='#1B1829', edgecolor='#312C45', labelcolor='#FAF6F0')
plt.title(f"Overall Security Risk Distribution\n({len(df)} Businesses Checked)", fontsize=16, weight="bold", pad=20, color="#E5A88B")
plt.savefig("charts/risk_distribution.png", dpi=300, bbox_inches='tight', facecolor='#120F1D')
plt.close()
print("Generating category scores bar chart...")
cat_scores = df.groupby("category")["security_score"].mean().sort_values(ascending=True)
fig, ax = plt.subplots(figsize=(8, 4.5))
palette = ["#F07167", "#E67E22", "#F1C40F", "#D5C3C6", "#E5A88B"] 
bars = ax.barh(cat_scores.index, cat_scores.values, color=palette, height=0.55, edgecolor='#312C45', alpha=0.95)
plt.xlabel("Average Security Score (0 - 5)", fontsize=12, weight="bold", labelpad=10)
plt.ylabel("Business Category", fontsize=12, weight="bold", labelpad=10)
plt.title("Average Security Score by Industry Category", fontsize=14, weight="bold", pad=15, color="#E5A88B")
plt.xlim(0, 5.5)
ax.grid(True, axis='x', linestyle='--', color='#312C45')
ax.tick_params(labelsize=11.5)
for bar in bars:
    width = bar.get_width()
    ax.text(width + 0.1, bar.get_y() + bar.get_height()/2, f'{width:.2f}', 
             va='center', ha='left', fontsize=12, weight='bold', color='#FAF6F0')
plt.tight_layout()
plt.savefig("charts/category_scores.png", dpi=300, facecolor='#120F1D')
plt.close()
print("Generating location scores bar chart...")
loc_scores = df.groupby("location")["security_score"].mean().sort_values(ascending=False)
fig, ax = plt.subplots(figsize=(8, 4.5))
bars = ax.bar(loc_scores.index, loc_scores.values, color="#E5A88B", width=0.45, edgecolor='#312C45', alpha=0.95) 
plt.xlabel("Neighborhood / Area in Delhi", fontsize=12, weight="bold", labelpad=10)
plt.ylabel("Average Security Score (0 - 5)", fontsize=12, weight="bold", labelpad=10)
plt.title("Geographical Security Analysis (Average Score)", fontsize=14, weight="bold", pad=15, color="#E5A88B")
plt.ylim(0, 5.5)
ax.grid(True, axis='y', linestyle='--', color='#312C45')
ax.tick_params(labelsize=11.5)
for bar in bars:
    height = bar.get_height()
    ax.text(bar.get_x() + bar.get_width()/2., height + 0.1, f'{height:.2f}', 
             ha='center', va='bottom', fontsize=11, weight='bold', color='#FAF6F0')
plt.xticks(rotation=15)
plt.tight_layout()
plt.savefig("charts/location_risks.png", dpi=300, facecolor='#120F1D')
plt.close()
print("Generating risk indicators frequency chart...")
indicators_flat = [indicator for sublist in df["risk_indicators"] for indicator in sublist]
indicator_counts = pd.Series(indicators_flat).value_counts().sort_values(ascending=True)
fig, ax = plt.subplots(figsize=(8, 4.5))
bars = ax.barh(indicator_counts.index, indicator_counts.values, color="#D5C3C6", height=0.55, edgecolor='#312C45', alpha=0.95) 
plt.xlabel("Number of Businesses Affected", fontsize=12, weight="bold", labelpad=10)
plt.ylabel("Security Vulnerability / Risk Indicator", fontsize=12, weight="bold", labelpad=10)
plt.title("Prevalence of Digital Risk Indicators", fontsize=14, weight="bold", pad=15, color="#E5A88B")
plt.xlim(0, max(indicator_counts.values) + 1.8)
ax.grid(True, axis='x', linestyle='--', color='#312C45')
ax.tick_params(labelsize=11.5)
for bar in bars:
    width = bar.get_width()
    ax.text(width + 0.15, bar.get_y() + bar.get_height()/2, f'{int(width)}', 
             va='center', ha='left', fontsize=12, weight='bold', color='#FAF6F0')
plt.tight_layout()
plt.savefig("charts/risk_indicators.png", dpi=300, facecolor='#120F1D')
plt.close()
print("Charts saved successfully in /charts folder.")
print("Generating Excel...")
wb = Workbook()
ws_dash = wb.active
ws_dash.title = "Executive Dashboard"
ws_data = wb.create_sheet(title="Security Assessment Data")
font_title = Font(name="Segoe UI", size=16, bold=True, color="231F33")
font_section = Font(name="Segoe UI", size=12, bold=True, color="231F33")
font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
font_body = Font(name="Segoe UI", size=10, color="2E2E2E")
font_body_bold = Font(name="Segoe UI", size=10, bold=True, color="2E2E2E")
font_kpi_num = Font(name="Segoe UI", size=20, bold=True, color="231F33")
font_kpi_label = Font(name="Segoe UI", size=9, color="7A7075", italic=True)
fill_header = PatternFill(start_color="231F33", end_color="231F33", fill_type="solid") 
fill_zebra = PatternFill(start_color="FAF7FA", end_color="FAF7FA", fill_type="solid") 
fill_kpi = PatternFill(start_color="F5EFEF", end_color="F5EFEF", fill_type="solid") 
fill_section_hdr = PatternFill(start_color="F9EBE5", end_color="F9EBE5", fill_type="solid") 
fills_risk = {
    "Very Secure": PatternFill(start_color="E8F8F0", end_color="E8F8F0", fill_type="solid"),
    "Secure": PatternFill(start_color="EDF7ED", end_color="EDF7ED", fill_type="solid"),
    "Moderate Risk": PatternFill(start_color="FFF9E6", end_color="FFF9E6", fill_type="solid"),
    "High Risk": PatternFill(start_color="FDF0E9", end_color="FDF0E9", fill_type="solid"),
    "Very High Risk": PatternFill(start_color="FDF1F0", end_color="FDF1F0", fill_type="solid")
}
fonts_risk = {
    "Very Secure": Font(name="Segoe UI", size=10, bold=True, color="1E6B43"),
    "Secure": Font(name="Segoe UI", size=10, bold=True, color="2E7D32"),
    "Moderate Risk": Font(name="Segoe UI", size=10, bold=True, color="B78103"),
    "High Risk": Font(name="Segoe UI", size=10, bold=True, color="A0522D"),
    "Very High Risk": Font(name="Segoe UI", size=10, bold=True, color="C0392B")
}
fills_priority = {
    "CRITICAL": PatternFill(start_color="FDF1F0", end_color="FDF1F0", fill_type="solid"),
    "MEDIUM": PatternFill(start_color="FFF9E6", end_color="FFF9E6", fill_type="solid"),
    "LOW": PatternFill(start_color="E8F8F0", end_color="E8F8F0", fill_type="solid")
}
fonts_priority = {
    "CRITICAL": Font(name="Segoe UI", size=10, bold=True, color="C0392B"),
    "MEDIUM": Font(name="Segoe UI", size=10, bold=True, color="B78103"),
    "LOW": Font(name="Segoe UI", size=10, bold=True, color="1E6B43")
}
thin_line = Side(style='thin', color='D9D9D9')
double_line = Side(style='double', color='000000')
border_all = Border(left=thin_line, right=thin_line, top=thin_line, bottom=thin_line)
border_header = Border(left=thin_line, right=thin_line, top=thin_line, bottom=Side(style='medium', color='1F4E78'))
border_total = Border(top=thin_line, bottom=double_line)
align_left = Alignment(horizontal='left', vertical='center')
align_center = Alignment(horizontal='center', vertical='center')
align_right = Alignment(horizontal='right', vertical='center')
align_left_wrap = Alignment(horizontal='left', vertical='center', wrap_text=True)
def create_dashboard_kpi(ws, col_start, col_end, formula_val, label_text, is_float=False):
    ws.merge_cells(f"{col_start}5:{col_end}5")
    ws.merge_cells(f"{col_start}6:{col_end}6")
    num_cell = ws[f"{col_start}5"]
    num_cell.value = formula_val
    num_cell.font = font_kpi_num
    num_cell.alignment = align_center
    num_cell.fill = fill_kpi
    if is_float:
        num_cell.number_format = '0.00'
    else:
        num_cell.number_format = '#,##0'
    lbl_cell = ws[f"{col_start}6"]
    lbl_cell.value = label_text
    lbl_cell.font = font_kpi_label
    lbl_cell.alignment = align_center
    lbl_cell.fill = fill_kpi
    for r in [5, 6]:
        col_ord_start = ord(col_start[0])
        col_ord_end = ord(col_end[0])
        for c_ord in range(col_ord_start, col_ord_end + 1):
            cell = ws[f"{chr(c_ord)}{r}"]
            cell.border = border_all
ws_data.views.sheetView[0].showGridLines = True
headers = [
    "Business ID", "Business Name", "Category", "Location", "Google Maps Rating", 
    "Reviews Count", "GBP Claimed", "Website Status", "Website URL", "HTTPS Redirect", 
    "SSL Status", "SSL Provider", "Admin Panel Exposed", "Forms Protection", 
    "Directory Listing", "Safe Browsing Check", "Missing Security Headers", "Security Score", 
    "Risk Level", "Outreach Priority", "Estimated Lead Value", "Outreach Strategy", 
    "Outreach Subject Line", "First Touch Pitch Hook", "Sales Pitch Hook", 
    "Key Security Gaps", "Actionable Recommendation"
]
for col_idx, h in enumerate(headers, 1):
    cell = ws_data.cell(row=2, column=col_idx, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = align_center
    cell.border = border_header
ws_data.row_dimensions[2].height = 30
data_start_row = 3
for r_idx, d in enumerate(processed_data, data_start_row):
    ws_data.cell(row=r_idx, column=1, value=d["id"]).alignment = align_center
    ws_data.cell(row=r_idx, column=2, value=d["name"]).alignment = align_left
    ws_data.cell(row=r_idx, column=3, value=d["category"]).alignment = align_left
    ws_data.cell(row=r_idx, column=4, value=d["location"]).alignment = align_center
    r_cell = ws_data.cell(row=r_idx, column=5, value=d["rating"])
    r_cell.alignment = align_right
    r_cell.number_format = '0.0'
    rev_cell = ws_data.cell(row=r_idx, column=6, value=d["reviews"])
    rev_cell.alignment = align_right
    rev_cell.number_format = '#,##0'
    ws_data.cell(row=r_idx, column=7, value=d["gbp_claimed"]).alignment = align_center
    ws_data.cell(row=r_idx, column=8, value=d["website_status"]).alignment = align_center
    ws_data.cell(row=r_idx, column=9, value=d["website_url"]).alignment = align_left
    ws_data.cell(row=r_idx, column=10, value=d["https_redirect"]).alignment = align_center
    ws_data.cell(row=r_idx, column=11, value=d["ssl_status"]).alignment = align_center
    ws_data.cell(row=r_idx, column=12, value=d["ssl_provider"]).alignment = align_center
    ws_data.cell(row=r_idx, column=13, value=d["admin_panel_exposed"]).alignment = align_center
    ws_data.cell(row=r_idx, column=14, value=d["forms_protection"]).alignment = align_center
    ws_data.cell(row=r_idx, column=15, value=d["directory_listing"]).alignment = align_center
    ws_data.cell(row=r_idx, column=16, value=d["safe_browsing"]).alignment = align_center
    ws_data.cell(row=r_idx, column=17, value=d["missing_headers"]).alignment = align_left
    score_cell = ws_data.cell(row=r_idx, column=18, value=d["security_score"])
    score_cell.alignment = align_center
    score_cell.number_format = '0'
    risk_cell = ws_data.cell(row=r_idx, column=19, value=d["risk_level"])
    risk_cell.alignment = align_center
    risk_cell.fill = fills_risk[d["risk_level"]]
    risk_cell.font = fonts_risk[d["risk_level"]]
    priority_cell = ws_data.cell(row=r_idx, column=20, value=d["outreach_priority"])
    priority_cell.alignment = align_center
    priority_cell.fill = fills_priority.get(d["outreach_priority"])
    priority_cell.font = fonts_priority.get(d["outreach_priority"])
    ws_data.cell(row=r_idx, column=21, value=d["estimated_lead_value"]).alignment = align_center
    ws_data.cell(row=r_idx, column=22, value=d["outreach_strategy"]).alignment = align_left_wrap
    ws_data.cell(row=r_idx, column=23, value=d["outreach_subject_line"]).alignment = align_left_wrap
    ws_data.cell(row=r_idx, column=24, value=d["first_touch_pitch"]).alignment = align_left_wrap
    ws_data.cell(row=r_idx, column=25, value=d["sales_pitch_hook"]).alignment = align_left_wrap
    ws_data.cell(row=r_idx, column=26, value=d["key_security_gaps"]).alignment = align_left_wrap
    ws_data.cell(row=r_idx, column=27, value=d["recommendation"]).alignment = align_left_wrap
    for col_idx in range(1, len(headers) + 1):
        c = ws_data.cell(row=r_idx, column=col_idx)
        if col_idx not in [19, 20]:
            c.font = font_body
            if r_idx % 2 == 0:
                c.fill = fill_zebra
        c.border = border_all
    max_lines = 1
    for col_idx in range(1, len(headers) + 1):
        cell_val = str(ws_data.cell(row=r_idx, column=col_idx).value or '')
        col_width = 12 
        if col_idx == 23: col_width = 45 
        elif col_idx == 22: col_width = 30 
        elif col_idx == 21: col_width = 35 
        elif col_idx == 17: col_width = 25 
        elif col_idx == 9: col_width = 30 
        lines = (len(cell_val) // col_width) + 1
        if lines > max_lines:
            max_lines = lines
    ws_data.row_dimensions[r_idx].height = max(18 * max_lines, 26)
data_end_row = data_start_row + len(processed_data) - 1
tot_row = data_end_row + 2
ws_data.cell(row=tot_row, column=1, value="Summary Metrics").font = font_body_bold
ws_data.cell(row=tot_row, column=1).alignment = align_left
ws_data.cell(row=tot_row, column=5, value=f"=AVERAGE(E3:E{data_end_row})").font = font_body_bold
ws_data.cell(row=tot_row, column=5).alignment = align_right
ws_data.cell(row=tot_row, column=5).number_format = '0.00'
ws_data.cell(row=tot_row, column=6, value=f"=SUM(F3:F{data_end_row})").font = font_body_bold
ws_data.cell(row=tot_row, column=6).alignment = align_right
ws_data.cell(row=tot_row, column=6).number_format = '#,##0'
ws_data.cell(row=tot_row, column=18, value=f"=AVERAGE(R3:R{data_end_row})").font = font_body_bold
ws_data.cell(row=tot_row, column=18).alignment = align_center
ws_data.cell(row=tot_row, column=18).number_format = '0.00'
for col_idx in range(1, len(headers) + 1):
    c = ws_data.cell(row=tot_row, column=col_idx)
    c.border = border_total
ws_data.row_dimensions[tot_row].height = 24
col_widths = {
    1: 12,  
    2: 25,  
    3: 22,  
    4: 15,  
    5: 15,  
    6: 14,  
    7: 14,  
    8: 15,  
    9: 25,  
    10: 15, 
    11: 15, 
    12: 18, 
    13: 18, 
    14: 18, 
    15: 18, 
    16: 18, 
    17: 22, 
    18: 14, 
    19: 16, 
    20: 16, 
    21: 20, 
    22: 25, 
    23: 30, 
    24: 35, 
    25: 35, 
    26: 30, 
    27: 40, 
}
for col_idx, width in col_widths.items():
    col_letter = get_column_letter(col_idx)
    ws_data.column_dimensions[col_letter].width = width
ws_dash.views.sheetView[0].showGridLines = True
ws_dash.merge_cells("A2:H2")
ws_dash["A2"] = "BUSINESS SECURITY & DIGITAL RISK ANALYSIS - PAN-INDIA"
ws_dash["A2"].font = Font(name="Segoe UI", size=18, bold=True, color="FFFFFF")
ws_dash["A2"].fill = PatternFill(start_color="0F2027", end_color="0F2027", fill_type="solid")
ws_dash["A2"].alignment = Alignment(horizontal='center', vertical='center')
ws_dash.row_dimensions[2].height = 40
ws_dash.merge_cells("A3:H3")
ws_dash["A3"] = "Digital Security Posture Assessment of 50 Businesses Across India"
ws_dash["A3"].font = Font(name="Segoe UI", size=11, italic=True, color="CCCCCC")
ws_dash["A3"].fill = PatternFill(start_color="203A43", end_color="203A43", fill_type="solid")
ws_dash["A3"].alignment = Alignment(horizontal='center', vertical='center')
ws_dash.row_dimensions[3].height = 22
risk_rows = [
    ("Very Secure", f"=COUNTIF('Security Assessment Data'!S3:S{data_end_row}, \"Very Secure\")"),
    ("Secure", f"=COUNTIF('Security Assessment Data'!S3:S{data_end_row}, \"Secure\")"),
    ("Moderate Risk", f"=COUNTIF('Security Assessment Data'!S3:S{data_end_row}, \"Moderate Risk\")"),
    ("High Risk", f"=COUNTIF('Security Assessment Data'!S3:S{data_end_row}, \"High Risk\")"),
    ("Very High Risk", f"=COUNTIF('Security Assessment Data'!S3:S{data_end_row}, \"Very High Risk\")")
]
cat_rows = [
    "Restaurants & Cafes",
    "Healthcare & Clinics",
    "Retail & Shops",
    "Professional Services",
    "Hospitality & Hotels"
]
total_checked = len(processed_data)
no_website_count = sum(1 for d in processed_data if d["website_status"] == "No Website")
exposed_admin_count = sum(1 for d in processed_data if d["admin_panel_exposed"] == "Yes")
unclaimed_count = sum(1 for d in processed_data if d["gbp_claimed"] == "No")
active_site_count = total_checked - no_website_count
missing_headers_count = sum(1 for d in processed_data if d["website_status"] == "Active" and d["missing_headers"] != "None")
missing_headers_pct = (missing_headers_count / active_site_count) * 100 if active_site_count > 0 else 0
observations = [
    "1. High Vulnerability in Professional Services: Professional services have a low average security score. Several regional law offices and tax consultants exhibit critical weaknesses.",
    f"2. Digital Absence Risk: {no_website_count} out of {total_checked} businesses analyzed have NO website, leaving them with a Very High Risk rating due to missed opportunities.",
    f"3. Missing Security Headers: {missing_headers_pct:.0f}% of the websites scanned are missing critical security headers like Content-Security-Policy (CSP) and HSTS.",
    f"4. Exposed Administrative Interfaces: {exposed_admin_count} businesses have exposed admin panels (/wp-admin or similar) accessible to the public.",
    f"5. Unclaimed Profiles: {unclaimed_count} businesses have unclaimed Google Business Profiles on Maps. This exposes them to direct profile hijacking and sabotage."
]
create_dashboard_kpi(ws_dash, "A", "B", f"=COUNTA('Security Assessment Data'!A3:A{data_end_row})", "Total Businesses Checked")
create_dashboard_kpi(ws_dash, "C", "D", f"=AVERAGE('Security Assessment Data'!R3:R{data_end_row})", "Average Security Score (0-5)", is_float=True)
create_dashboard_kpi(ws_dash, "E", "F", f"=COUNTIF('Security Assessment Data'!S3:S{data_end_row}, \"*Secure*\")", "Secure Businesses (Score 4-5)")
create_dashboard_kpi(ws_dash, "G", "H", f"=COUNTIF('Security Assessment Data'!S3:S{data_end_row}, \"*Risk*\")", "Risky Businesses (Score 0-3)")
ws_dash.row_dimensions[5].height = 32
ws_dash.row_dimensions[6].height = 18
ws_dash.row_dimensions[9].height = 24
ws_dash["A9"] = "Risk Level Breakdown"
ws_dash["A9"].font = font_section
ws_dash["A9"].fill = fill_section_hdr
ws_dash["A9"].alignment = align_left
ws_dash.merge_cells("A9:C9")
ws_dash["A9"].border = border_all
ws_dash["B9"].border = border_all
ws_dash["C9"].border = border_all
risk_headers = ["Risk Level", "Count", "Percentage"]
for idx, h in enumerate(risk_headers, 1):
    c = ws_dash.cell(row=10, column=idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_center
    c.border = border_header
ws_dash.row_dimensions[10].height = 24
for idx, (level, formula) in enumerate(risk_rows, 11):
    c_lvl = ws_dash.cell(row=idx, column=1, value=level)
    c_lvl.font = font_body
    c_lvl.border = border_all
    c_lvl.alignment = align_left
    c_cnt = ws_dash.cell(row=idx, column=2, value=formula)
    c_cnt.font = font_body
    c_cnt.border = border_all
    c_cnt.alignment = align_center
    c_pct = ws_dash.cell(row=idx, column=3, value=f"=B{idx}/SUM(B11:B15)")
    c_pct.font = font_body
    c_pct.border = border_all
    c_pct.alignment = align_right
    c_pct.number_format = '0.0%'
    ws_dash.row_dimensions[idx].height = 20
tot_level_row = 16
ws_dash.cell(row=tot_level_row, column=1, value="Total").font = font_body_bold
ws_dash.cell(row=tot_level_row, column=1).alignment = align_left
ws_dash.cell(row=tot_level_row, column=2, value="=SUM(B11:B15)").font = font_body_bold
ws_dash.cell(row=tot_level_row, column=2).alignment = align_center
ws_dash.cell(row=tot_level_row, column=3, value="=SUM(C11:C15)").font = font_body_bold
ws_dash.cell(row=tot_level_row, column=3).alignment = align_right
ws_dash.cell(row=tot_level_row, column=3).number_format = '0.0%'
for c_idx in range(1, 4):
    ws_dash.cell(row=tot_level_row, column=c_idx).border = border_total
ws_dash.column_dimensions['D'].width = 4
ws_dash["E9"] = "Average Score by Category"
ws_dash["E9"].font = font_section
ws_dash["E9"].fill = fill_section_hdr
ws_dash["E9"].alignment = align_left
ws_dash.merge_cells("E9:G9")
ws_dash["E9"].border = border_all
ws_dash["F9"].border = border_all
ws_dash["G9"].border = border_all
cat_headers = ["Category", "Total Checked", "Average Score"]
for idx, h in enumerate(cat_headers, 5):
    c = ws_dash.cell(row=10, column=idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = align_center
    c.border = border_header
for idx, cat in enumerate(cat_rows, 11):
    c_cat = ws_dash.cell(row=idx, column=5, value=cat)
    c_cat.font = font_body
    c_cat.border = border_all
    c_cat.alignment = align_left
    c_cnt = ws_dash.cell(row=idx, column=6, value=f"=COUNTIF('Security Assessment Data'!C3:C{data_end_row}, \"{cat}\")")
    c_cnt.font = font_body
    c_cnt.border = border_all
    c_cnt.alignment = align_center
    c_avg = ws_dash.cell(row=idx, column=7, value=f"=AVERAGEIF('Security Assessment Data'!C3:C{data_end_row}, \"{cat}\", 'Security Assessment Data'!R3:R{data_end_row})")
    c_avg.font = font_body
    c_avg.border = border_all
    c_avg.alignment = align_center
    c_avg.number_format = '0.00'
    ws_dash.row_dimensions[idx].height = 20
tot_cat_row = 16
ws_dash.cell(row=tot_cat_row, column=5, value="All Industries").font = font_body_bold
ws_dash.cell(row=tot_cat_row, column=5).alignment = align_left
ws_dash.cell(row=tot_cat_row, column=6, value="=SUM(F11:F15)").font = font_body_bold
ws_dash.cell(row=tot_cat_row, column=6).alignment = align_center
ws_dash.cell(row=tot_cat_row, column=7, value="=AVERAGE('Security Assessment Data'!R3:R{data_end_row})").font = font_body_bold
ws_dash.cell(row=tot_cat_row, column=7).alignment = align_center
ws_dash.cell(row=tot_cat_row, column=7).number_format = '0.00'
for c_idx in range(5, 8):
    ws_dash.cell(row=tot_cat_row, column=c_idx).border = border_total
ws_dash["A19"] = "Key Observation & Analysis Insights:"
ws_dash["A19"].font = font_section
ws_dash["A19"].fill = fill_section_hdr
ws_dash.merge_cells("A19:H19")
for c_ord in range(ord('A'), ord('H') + 1):
    ws_dash[f"{chr(c_ord)}19"].border = border_all
ws_dash.row_dimensions[19].height = 24
for idx, obs in enumerate(observations, 20):
    ws_dash.cell(row=idx, column=1, value=obs).font = Font(name="Segoe UI", size=10, color="555555")
    ws_dash.cell(row=idx, column=1).alignment = align_left
    ws_dash.merge_cells(start_row=idx, start_column=1, end_row=idx, end_column=8)
    ws_dash.row_dimensions[idx].height = 20
ws_dash.column_dimensions['A'].width = 24
ws_dash.column_dimensions['B'].width = 12
ws_dash.column_dimensions['C'].width = 12
ws_dash.column_dimensions['D'].width = 4
ws_dash.column_dimensions['E'].width = 26
ws_dash.column_dimensions['F'].width = 15
ws_dash.column_dimensions['G'].width = 15
ws_dash.column_dimensions['H'].width = 12
wb.save("National_Digital_Risk_Audit_Database.xlsx")
print("Excel Workbook National_Digital_Risk_Audit_Database.xlsx generated and styled successfully!")
print("Generating PPTX presentation...")
prs = Presentation()
prs.slide_width = Inches(16.0) 
prs.slide_height = Inches(9.0)
DARK_BG = RGBColor(18, 15, 29)        
CARD_BG = RGBColor(27, 24, 41)        
BORDER_COLOR = RGBColor(49, 44, 69)   
TEXT_WHITE = RGBColor(250, 246, 240)  
TEXT_MUTED = RGBColor(213, 195, 198)  
CYAN_ACCENT = RGBColor(229, 168, 139)  
GREEN_SECURE = RGBColor(136, 212, 158)  
RED_RISK = RGBColor(240, 113, 103)    
YELLOW_WARN = RGBColor(241, 196, 15)   
def shorten_gaps(gaps):
    gaps = gaps.replace("Exposed Admin Panel (/wp-admin or /admin)", "Exposed Admin Page")
    gaps = gaps.replace("Digital Absence, Unclaimed Google Business Profile", "No Website, Unclaimed Profile")
    gaps = gaps.replace("Digital Absence, No Website Present", "No Website (Digital Absence)")
    gaps = gaps.replace("No HTTPS Redirect", "HTTP Only (No HTTPS)")
    gaps = gaps.replace("Unprotected Forms, Directory Listing Enabled", "Unprotected Forms, Open Dir")
    gaps = gaps.replace("All Headers Missing", "Missing Headers")
    gaps = gaps.replace("SSL Status: Expired", "Expired SSL")
    gaps = gaps.replace("SSL Status: None", "No SSL")
    return gaps
def apply_background(slide, color=DARK_BG):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color
def create_title(slide, text, subtitle_text=None, top_inches=0.5):
    title_box = slide.shapes.add_textbox(Inches(1.0), Inches(top_inches), Inches(14.0), Inches(1.1))
    tf = title_box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = "Segoe UI"
    p.font.size = Pt(44) 
    p.font.bold = True
    p.font.color.rgb = TEXT_WHITE
    if subtitle_text:
        p2 = tf.add_paragraph()
        p2.text = subtitle_text
        p2.font.name = "Segoe UI"
        p2.font.size = Pt(18) 
        p2.font.color.rgb = CYAN_ACCENT
        p2.space_before = Pt(4)
    badge_box = slide.shapes.add_textbox(Inches(11.0), Inches(top_inches), Inches(4.0), Inches(0.5))
    tf_b = badge_box.text_frame
    tf_b.word_wrap = False
    tf_b.margin_left = tf_b.margin_top = tf_b.margin_right = tf_b.margin_bottom = 0
    p_b = tf_b.paragraphs[0]
    p_b.text = "PAN-INDIA DIGITAL SECURITY AUDIT • 2026"
    p_b.alignment = PP_ALIGN.RIGHT
    p_b.font.name = "Segoe UI"
    p_b.font.size = Pt(10)
    p_b.font.bold = True
    p_b.font.color.rgb = TEXT_MUTED
    divider = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, 
        Inches(1.0), Inches(top_inches + 1.2), Inches(14.0), Inches(0.05)
    )
    divider.fill.solid()
    divider.fill.fore_color.rgb = CYAN_ACCENT
    divider.line.color.rgb = CYAN_ACCENT
    divider.line.width = Pt(1)
def draw_premium_card(slide, x, y, w, h, bg_color=CARD_BG, border_color=BORDER_COLOR, accent_color=CYAN_ACCENT):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, 
        Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.color.rgb = border_color
    shape.line.width = Pt(1)
    shape.adjustments[0] = 0.04
    if accent_color:
        accent = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(x), Inches(y), Inches(w), Inches(0.08)
        )
        accent.fill.solid()
        accent.fill.fore_color.rgb = accent_color
        accent.line.color.rgb = accent_color
        accent.line.width = Pt(1)
    return shape
slide_layout = prs.slide_layouts[6] 
slide1 = prs.slides.add_slide(slide_layout)
apply_background(slide1)
draw_premium_card(slide1, 1.0, 2.0, 14.0, 5.0, bg_color=CARD_BG, border_color=BORDER_COLOR, accent_color=CYAN_ACCENT)
tb = slide1.shapes.add_textbox(Inches(1.5), Inches(2.5), Inches(13.0), Inches(4.0))
tf = tb.text_frame
tf.word_wrap = True
p_sub = tf.paragraphs[0]
p_sub.text = "WEEK 3 - CYBER SECURITY TASK"
p_sub.font.name = "Segoe UI"
p_sub.font.size = Pt(20) 
p_sub.font.bold = True
p_sub.font.color.rgb = CYAN_ACCENT
p_main = tf.add_paragraph()
p_main.text = "BUSINESS SECURITY & DIGITAL RISK ANALYSIS"
p_main.font.name = "Segoe UI"
p_main.font.size = Pt(44) 
p_main.font.bold = True
p_main.font.color.rgb = TEXT_WHITE
p_main.space_before = Pt(12)
p_desc = tf.add_paragraph()
p_desc.text = "Digital presence assessment, vulnerability scan, and cyber risk scoring of 50 businesses identified via Google Maps across major Indian cities."
p_desc.font.name = "Segoe UI"
p_desc.font.size = Pt(16) 
p_desc.font.color.rgb = TEXT_MUTED
p_desc.space_before = Pt(20)
p_author = tf.add_paragraph()
p_author.text = "Prepared by: Cyber Security Consultant  |  Assessment Scope: Pan-India Market Audit"
p_author.font.name = "Segoe UI"
p_author.font.size = Pt(14) 
p_author.font.color.rgb = CYAN_ACCENT
p_author.space_before = Pt(50)
slide2 = prs.slides.add_slide(slide_layout)
apply_background(slide2)
create_title(slide2, "Project Scope & Assessment Areas", "Target Region: India Local Market Analysis (50 Businesses)")
col_w = 4.4
gap = 0.4
y_pos = 2.4
h_pos = 5.6
draw_premium_card(slide2, 1.0, y_pos, col_w, h_pos, accent_color=CYAN_ACCENT)
tb_sc1 = slide2.shapes.add_textbox(Inches(1.2), Inches(y_pos + 0.3), Inches(col_w - 0.4), Inches(h_pos - 0.6))
tf = tb_sc1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "01  |  Data Sourcing"
p.font.name = "Segoe UI"
p.font.size = Pt(26) 
p.font.bold = True
p.font.color.rgb = CYAN_ACCENT
p.space_after = Pt(16)
p_body1_hdr = tf.add_paragraph()
p_body1_hdr.text = "Google Maps Sourcing"
p_body1_hdr.font.name = "Segoe UI"
p_body1_hdr.font.size = Pt(20) 
p_body1_hdr.font.bold = True
p_body1_hdr.font.color.rgb = TEXT_WHITE
p_body1 = tf.add_paragraph()
p_body1.text = "• Mapped 50 Indian businesses.\n• Mumbai, Bangalore, Chennai, Kolkata.\n• Delhi, Hyderabad, Pune, Jaipur, Goa."
p_body1.font.name = "Segoe UI"
p_body1.font.size = Pt(18) 
p_body1.font.color.rgb = TEXT_MUTED
p_body1.space_before = Pt(4)
p_body1.space_after = Pt(16)
p_body2_hdr = tf.add_paragraph()
p_body2_hdr.text = "Security Audits"
p_body2_hdr.font.name = "Segoe UI"
p_body2_hdr.font.size = Pt(20) 
p_body2_hdr.font.bold = True
p_body2_hdr.font.color.rgb = TEXT_WHITE
p_body2 = tf.add_paragraph()
p_body2.text = "• Scanned target websites.\n• Checked SSL & admin page flags."
p_body2.font.name = "Segoe UI"
p_body2.font.size = Pt(18) 
p_body2.font.color.rgb = TEXT_MUTED
p_body2.space_before = Pt(4)
draw_premium_card(slide2, 1.0 + col_w + gap, y_pos, col_w, h_pos, accent_color=CYAN_ACCENT)
tb_sc2 = slide2.shapes.add_textbox(Inches(1.2 + col_w + gap), Inches(y_pos + 0.3), Inches(col_w - 0.4), Inches(h_pos - 0.6))
tf = tb_sc2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "02  |  Industry Scope"
p.font.name = "Segoe UI"
p.font.size = Pt(26)
p.font.bold = True
p.font.color.rgb = CYAN_ACCENT
p.space_after = Pt(16)
p_body1_hdr = tf.add_paragraph()
p_body1_hdr.text = "Sectors Evaluated"
p_body1_hdr.font.name = "Segoe UI"
p_body1_hdr.font.size = Pt(20)
p_body1_hdr.font.bold = True
p_body1_hdr.font.color.rgb = TEXT_WHITE
p_body1 = tf.add_paragraph()
p_body1.text = "• Restaurants & Cafes (10 checked)\n• Healthcare & Clinics (10 checked)\n• Retail & Shops (10 checked)\n• Professional Services (10 checked)\n• Hospitality & Hotels (10 checked)"
p_body1.font.name = "Segoe UI"
p_body1.font.size = Pt(18) 
p_body1.font.color.rgb = TEXT_MUTED
p_body1.space_before = Pt(4)
p_body1.space_after = Pt(16)
draw_premium_card(slide2, 1.0 + 2*(col_w + gap), y_pos, col_w, h_pos, accent_color=CYAN_ACCENT)
tb_sc3 = slide2.shapes.add_textbox(Inches(1.2 + 2*(col_w + gap)), Inches(y_pos + 0.3), Inches(col_w - 0.4), Inches(h_pos - 0.6))
tf = tb_sc3.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "03  |  Security Parameters"
p.font.name = "Segoe UI"
p.font.size = Pt(26)
p.font.bold = True
p.font.color.rgb = CYAN_ACCENT
p.space_after = Pt(16)
p_body1_hdr = tf.add_paragraph()
p_body1_hdr.text = "Protocols & SSL Checks"
p_body1_hdr.font.name = "Segoe UI"
p_body1_hdr.font.size = Pt(20)
p_body1_hdr.font.bold = True
p_body1_hdr.font.color.rgb = TEXT_WHITE
p_body1 = tf.add_paragraph()
p_body1.text = "• Checked HTTPS redirect.\n• Checked SSL certificate validity.\n• Scanned admin login exposure."
p_body1.font.name = "Segoe UI"
p_body1.font.size = Pt(18) 
p_body1.font.color.rgb = TEXT_MUTED
p_body1.space_before = Pt(4)
p_body1.space_after = Pt(16)
p_body2_hdr = tf.add_paragraph()
p_body2_hdr.text = "Asset Exposures"
p_body2_hdr.font.name = "Segoe UI"
p_body2_hdr.font.size = Pt(20)
p_body2_hdr.font.bold = True
p_body2_hdr.font.color.rgb = TEXT_WHITE
p_body2 = tf.add_paragraph()
p_body2.text = "• Scanned open directory indexes.\n• Tested contact forms protection."
p_body2.font.name = "Segoe UI"
p_body2.font.size = Pt(18) 
p_body2.font.color.rgb = TEXT_MUTED
p_body2.space_before = Pt(4)
slide3 = prs.slides.add_slide(slide_layout)
apply_background(slide3)
create_title(slide3, "Executive Summary", "Key Findings and Performance Metrics")
kpi_w = 3.2
kpi_gap = 0.4
y_kpi = 2.4
h_kpi = 1.8
total_checked = len(df)
avg_score = df["security_score"].mean()
secure_count = len(df[df["security_score"] >= 4])
secure_rate = (secure_count / total_checked) * 100
risk_rate = 100 - secure_rate
no_website_count = len(df[df["website_status"] == "No Website"])
unclaimed_count = len(df[df["gbp_claimed"] == "No"])
exposed_admin_count = len(df[df["admin_panel_exposed"] == "Yes"])
kpi_data = [
    {"num": str(total_checked), "lbl": "TOTAL CHECKED", "color": TEXT_WHITE, "accent": TEXT_MUTED},
    {"num": f"{avg_score:.2f}", "lbl": "AVERAGE SCORE", "color": CYAN_ACCENT, "accent": CYAN_ACCENT},
    {"num": f"{secure_rate:.0f}%", "lbl": "SECURE RATE", "color": GREEN_SECURE, "accent": GREEN_SECURE},
    {"num": f"{risk_rate:.0f}%", "lbl": "AT RISK RATE", "color": RED_RISK, "accent": RED_RISK}
]
for idx, kd in enumerate(kpi_data):
    x_k = 1.0 + idx * (kpi_w + kpi_gap)
    draw_premium_card(slide3, x_k, y_kpi, kpi_w, h_kpi, accent_color=kd["accent"])
    tb_k = slide3.shapes.add_textbox(Inches(x_k + 0.1), Inches(y_kpi + 0.1), Inches(kpi_w - 0.2), Inches(h_kpi - 0.2))
    tf = tb_k.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p_num = tf.paragraphs[0]
    p_num.text = kd["num"]
    p_num.alignment = PP_ALIGN.CENTER
    p_num.font.name = "Segoe UI"
    p_num.font.size = Pt(48) 
    p_num.font.bold = True
    p_num.font.color.rgb = kd["color"]
    p_lbl = tf.add_paragraph()
    p_lbl.text = kd["lbl"]
    p_lbl.alignment = PP_ALIGN.CENTER
    p_lbl.font.name = "Segoe UI"
    p_lbl.font.size = Pt(15) 
    p_lbl.font.bold = True
    p_lbl.font.color.rgb = TEXT_MUTED
    p_lbl.space_before = Pt(3)
draw_premium_card(slide3, 1.0, 4.6, 14.0, 3.4, accent_color=CYAN_ACCENT)
tb_summary = slide3.shapes.add_textbox(Inches(1.3), Inches(4.8), Inches(13.4), Inches(3.0))
tf = tb_summary.text_frame
tf.word_wrap = True
p_title = tf.paragraphs[0]
p_title.text = "High-Level Cybersecurity Posture Analysis"
p_title.font.name = "Segoe UI"
p_title.font.size = Pt(26) 
p_title.font.bold = True
p_title.font.color.rgb = CYAN_ACCENT
p_title.space_after = Pt(10)
p_text = tf.add_paragraph()
p_text.text = (
    f"• Security Divide: {secure_rate:.0f}% Secure rate is driven by corporate hotels/hospitals, but {risk_rate:.0f}% At-Risk rate exposes local small businesses.\n"
    f"• Protocol Gaps: Plaintext transmissions and exposed CMS panels ({exposed_admin_count} sites) are the largest source of website vulnerability.\n"
    f"• Sabotage Risks: {unclaimed_count} Google Maps business profiles remain unclaimed, presenting an active hijacking risk for information edits."
)
p_text.font.name = "Segoe UI"
p_text.font.size = Pt(18) 
p_text.font.color.rgb = TEXT_WHITE
p_text.space_before = Pt(4)
slide4 = prs.slides.add_slide(slide_layout)
apply_background(slide4)
create_title(slide4, "Security Risk & Score Distribution", "Proportion of overall secure vs. vulnerable businesses")
draw_premium_card(slide4, 1.0, 2.4, 6.5, 5.6, accent_color=CYAN_ACCENT)
tb_dist = slide4.shapes.add_textbox(Inches(1.3), Inches(2.6), Inches(5.9), Inches(5.2))
tf = tb_dist.text_frame
tf.word_wrap = True
p_h = tf.paragraphs[0]
p_h.text = "Risk Tier Analysis"
p_h.font.name = "Segoe UI"
p_h.font.size = Pt(26) 
p_h.font.bold = True
p_h.font.color.rgb = CYAN_ACCENT
p_h.space_after = Pt(14)
p1_title = tf.add_paragraph()
p1_title.text = f"Very Secure (Score 5) — {len(df[df['security_score'] == 5])} Businesses"
p1_title.font.name = "Segoe UI"
p1_title.font.bold = True
p1_title.font.size = Pt(20) 
p1_title.font.color.rgb = GREEN_SECURE
p1_desc = tf.add_paragraph()
p1_desc.text = "Fully secure websites with active HTTPS redirections, valid SSL certificates, and hidden CMS login pages."
p1_desc.font.name = "Segoe UI"
p1_desc.font.size = Pt(18) 
p1_desc.font.color.rgb = TEXT_WHITE
p1_desc.space_after = Pt(12)
p2_title = tf.add_paragraph()
p2_title.text = f"Secure (Score 4) — {len(df[df['security_score'] == 4])} Businesses"
p2_title.font.name = "Segoe UI"
p2_title.font.bold = True
p2_title.font.size = Pt(20) 
p2_title.font.color.rgb = CYAN_ACCENT
p2_desc = tf.add_paragraph()
p2_desc.text = "Minor vulnerabilities detected (missing CSP or HSTS headers), but user input forms are secure."
p2_desc.font.name = "Segoe UI"
p2_desc.font.size = Pt(18) 
p2_desc.font.color.rgb = TEXT_WHITE
p2_desc.space_after = Pt(12)
p3_title = tf.add_paragraph()
p3_title.text = f"Vulnerable (Score 0-3) — {len(df[df['security_score'] <= 3])} Businesses"
p3_title.font.name = "Segoe UI"
p3_title.font.bold = True
p3_title.font.size = Pt(20) 
p3_title.font.color.rgb = RED_RISK
p3_desc = tf.add_paragraph()
p3_desc.text = "Critical protocol exposures or complete digital absence (no website) with unclaimed business profiles."
p3_desc.font.name = "Segoe UI"
p3_desc.font.size = Pt(18) 
p3_desc.font.color.rgb = TEXT_WHITE
chart_path_1 = "charts/risk_distribution.png"
if os.path.exists(chart_path_1):
    slide4.shapes.add_picture(chart_path_1, Inches(8.0), Inches(2.4), width=Inches(7.0), height=Inches(5.6))
slide5 = prs.slides.add_slide(slide_layout)
apply_background(slide5)
create_title(slide5, "Vulnerability & Risk Indicators", "Prevalence of specific digital risks across the assessed group")
chart_path_2 = "charts/risk_indicators.png"
if os.path.exists(chart_path_2):
    slide5.shapes.add_picture(chart_path_2, Inches(1.0), Inches(2.4), width=Inches(7.0), height=Inches(5.6))
draw_premium_card(slide5, 8.5, 2.4, 6.5, 5.6, accent_color=RED_RISK)
tb_vul = slide5.shapes.add_textbox(Inches(8.8), Inches(2.6), Inches(5.9), Inches(5.2))
tf = tb_vul.text_frame
tf.word_wrap = True
p_h = tf.paragraphs[0]
p_h.text = "Vulnerability Breakdown"
p_h.font.name = "Segoe UI"
p_h.font.size = Pt(26) 
p_h.font.bold = True
p_h.font.color.rgb = CYAN_ACCENT
p_h.space_after = Pt(14)
p1_title = tf.add_paragraph()
insecure_count = len(df[(df["website_status"] == "Active") & ((df["ssl_status"] == "None") | (df["ssl_status"] == "Expired") | (df["https_redirect"] == "No"))])
p1_title.text = f"Insecure Protocols ({insecure_count} Sites)"
p1_title.font.name = "Segoe UI"
p1_title.font.bold = True
p1_title.font.size = Pt(20) 
p1_title.font.color.rgb = RED_RISK
p1_desc = tf.add_paragraph()
p1_desc.text = "Transmitting booking and user details in plaintext over HTTP. Vulnerable to interception."
p1_desc.font.name = "Segoe UI"
p1_desc.font.size = Pt(18) 
p1_desc.font.color.rgb = TEXT_WHITE
p1_desc.space_after = Pt(12)
p2_title = tf.add_paragraph()
exposed_admin_count = len(df[df["admin_panel_exposed"] == "Yes"])
p2_title.text = f"Exposed Admin Panels ({exposed_admin_count} Sites)"
p2_title.font.name = "Segoe UI"
p2_title.font.bold = True
p2_title.font.size = Pt(20) 
p2_title.font.color.rgb = RED_RISK
p2_desc = tf.add_paragraph()
p2_desc.text = "WordPress login directories are exposed, allowing hackers to run automated brute-force scans."
p2_desc.font.name = "Segoe UI"
p2_desc.font.size = Pt(18) 
p2_desc.font.color.rgb = TEXT_WHITE
p2_desc.space_after = Pt(12)
p3_title = tf.add_paragraph()
open_dir_count = len(df[df["directory_listing"] == "Yes"])
p3_title.text = f"Open Directories ({open_dir_count} Sites)"
p3_title.font.name = "Segoe UI"
p3_title.font.bold = True
p3_title.font.size = Pt(20) 
p3_title.font.color.rgb = RED_RISK
p3_desc = tf.add_paragraph()
p3_desc.text = "Folder indexing is active, exposing backend assets and file directories to public view."
p3_desc.font.name = "Segoe UI"
p3_desc.font.size = Pt(18) 
p3_desc.font.color.rgb = TEXT_WHITE
slide6 = prs.slides.add_slide(slide_layout)
apply_background(slide6)
create_title(slide6, "Industry Category & Location Analysis", "Comparing security standards across sectors and areas")
draw_premium_card(slide6, 1.0, 2.4, 6.5, 5.6, accent_color=CYAN_ACCENT)
tb_loc = slide6.shapes.add_textbox(Inches(1.3), Inches(2.6), Inches(5.9), Inches(5.2))
tf = tb_loc.text_frame
tf.word_wrap = True
p_h = tf.paragraphs[0]
p_h.text = "Industry & Geography Risks"
p_h.font.name = "Segoe UI"
p_h.font.size = Pt(26) 
p_h.font.bold = True
p_h.font.color.rgb = CYAN_ACCENT
p_h.space_after = Pt(14)
p1_title = tf.add_paragraph()
p1_title.text = "Industry Security Scorecard"
p1_title.font.name = "Segoe UI"
p1_title.font.bold = True
p1_title.font.size = Pt(20) 
p1_title.font.color.rgb = CYAN_ACCENT
p1_desc = tf.add_paragraph()
p1_desc.text = "Hospitality and Healthcare lead in security standards due to compliance regulations. Professional Services and Retail lag behind due to high rates of digital absence and unencrypted sites."
p1_desc.font.name = "Segoe UI"
p1_desc.font.size = Pt(18) 
p1_desc.font.color.rgb = TEXT_WHITE
p1_desc.space_after = Pt(14)
p2_title = tf.add_paragraph()
p2_title.text = "Geographical Analysis"
p2_title.font.name = "Segoe UI"
p2_title.font.bold = True
p2_title.font.size = Pt(20) 
p2_title.font.color.rgb = CYAN_ACCENT
p2_desc = tf.add_paragraph()
p2_desc.text = "Major metropolitan centers show stark contrasts: premium corporate entities in Mumbai and Bengaluru exhibit high security postures, while smaller regional firms represent high-vulnerability clusters."
p2_desc.font.name = "Segoe UI"
p2_desc.font.size = Pt(18) 
p2_desc.font.color.rgb = TEXT_WHITE
chart_path_3 = "charts/category_scores.png"
if os.path.exists(chart_path_3):
    slide6.shapes.add_picture(chart_path_3, Inches(8.0), Inches(2.4), width=Inches(7.0), height=Inches(5.6))
slide7 = prs.slides.add_slide(slide_layout)
apply_background(slide7)
create_title(slide7, "Top 10 Risky Businesses", "Identified targets requiring immediate intervention")
df_sorted = df.sort_values(by=["security_score", "reviews"], ascending=[True, False]).head(10)
rows, cols = 11, 5
left, top, width, height = Inches(1.0), Inches(2.4), Inches(14.0), Inches(5.8)
table_shape = slide7.shapes.add_table(rows, cols, left, top, width, height)
table = table_shape.table
table.rows[0].height = Inches(0.60)
for r in range(1, 11):
    table.rows[r].height = Inches(0.65) 
table.columns[0].width = Inches(3.3)  
table.columns[1].width = Inches(2.4)  
table.columns[2].width = Inches(1.8)  
table.columns[3].width = Inches(1.5)  
table.columns[4].width = Inches(5.0) 
headers_tbl = ["Business Name", "Category", "Location", "Security Score", "Primary Security Gaps"]
for col_idx, text in enumerate(headers_tbl):
    cell = table.cell(0, col_idx)
    cell.text = text
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE 
    cell.fill.solid()
    cell.fill.fore_color.rgb = CARD_BG
    for p in cell.text_frame.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        p.font.name = "Segoe UI"
        p.font.size = Pt(16) 
        p.font.bold = True
        p.font.color.rgb = CYAN_ACCENT
for row_idx, (_, r) in enumerate(df_sorted.iterrows(), 1):
    table.cell(row_idx, 0).text = r["name"]
    table.cell(row_idx, 1).text = r["category"]
    table.cell(row_idx, 2).text = r["location"]
    table.cell(row_idx, 3).text = f"{int(r['security_score'])} / 5"
    table.cell(row_idx, 4).text = shorten_gaps(r["key_security_gaps"])
    for col_idx in range(cols):
        cell = table.cell(row_idx, col_idx)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE 
        cell.fill.solid()
        cell.fill.fore_color.rgb = DARK_BG
        score_cell = table.cell(row_idx, 3)
        for p in score_cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            p.font.bold = True
            if r["security_score"] <= 1:
                p.font.color.rgb = RED_RISK
            elif r["security_score"] == 2:
                p.font.color.rgb = RGBColor(230, 126, 34)
            else:
                p.font.color.rgb = RGBColor(241, 196, 15)
        for p in cell.text_frame.paragraphs:
            p.font.name = "Segoe UI"
            p.font.size = Pt(14) 
            if col_idx != 3:
                p.font.color.rgb = TEXT_WHITE
            if col_idx == 4:
                p.alignment = PP_ALIGN.LEFT
            elif col_idx in [1, 2]:
                p.alignment = PP_ALIGN.CENTER
slide8 = prs.slides.add_slide(slide_layout)
apply_background(slide8)
create_title(slide8, "Strategic Cybersecurity Roadmap", "Actionable recommendations for businesses and service providers")
w_q, h_q = 6.8, 2.7
x_l, x_r = 1.0, 8.2
y_t, y_b = 2.4, 5.5
draw_premium_card(slide8, x_l, y_t, w_q, h_q, accent_color=RED_RISK)
tb_q1 = slide8.shapes.add_textbox(Inches(x_l + 0.2), Inches(y_t + 0.2), Inches(w_q - 0.4), Inches(h_q - 0.4))
tf = tb_q1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "1. Immediate Remediation (At-Risk Sites)"
p.font.name = "Segoe UI"
p.font.size = Pt(20) 
p.font.bold = True
p.font.color.rgb = CYAN_ACCENT
p.space_after = Pt(8)
p_b1 = tf.add_paragraph()
p_b1.text = "• Install SSL Certificates: Force HTTP to HTTPS redirects."
p_b1.font.name = "Segoe UI"
p_b1.font.size = Pt(15) 
p_b1.font.color.rgb = TEXT_WHITE
p_b2 = tf.add_paragraph()
p_b2.text = "• Cloak Admin Panels: Restrict access to /wp-admin."
p_b2.font.name = "Segoe UI"
p_b2.font.size = Pt(15) 
p_b2.font.color.rgb = TEXT_WHITE
p_b3 = tf.add_paragraph()
p_b3.text = "• Secure Forms: Deploy anti-CSRF tokens on forms."
p_b3.font.name = "Segoe UI"
p_b3.font.size = Pt(15) 
p_b3.font.color.rgb = TEXT_WHITE
draw_premium_card(slide8, x_r, y_t, w_q, h_q, accent_color=CYAN_ACCENT)
tb_q2 = slide8.shapes.add_textbox(Inches(x_r + 0.2), Inches(y_t + 0.2), Inches(w_q - 0.4), Inches(h_q - 0.4))
tf = tb_q2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "2. Google Maps Profile Hardening"
p.font.name = "Segoe UI"
p.font.size = Pt(20) 
p.font.bold = True
p.font.color.rgb = CYAN_ACCENT
p.space_after = Pt(8)
p_b1 = tf.add_paragraph()
p_b1.text = "• Profile Claiming: Verify unclaimed profiles on Maps."
p_b1.font.name = "Segoe UI"
p_b1.font.size = Pt(15) 
p_b1.font.color.rgb = TEXT_WHITE
p_b2 = tf.add_paragraph()
p_b2.text = "• Access Audits: Enforce MFA on associated accounts."
p_b2.font.name = "Segoe UI"
p_b2.font.size = Pt(15) 
p_b2.font.color.rgb = TEXT_WHITE
p_b3 = tf.add_paragraph()
p_b3.text = "• Monitor Edits: Set alerts for Google Maps suggestions."
p_b3.font.name = "Segoe UI"
p_b3.font.size = Pt(15) 
p_b3.font.color.rgb = TEXT_WHITE
draw_premium_card(slide8, x_l, y_b, w_q, h_q, accent_color=GREEN_SECURE)
tb_q3 = slide8.shapes.add_textbox(Inches(x_l + 0.2), Inches(y_b + 0.2), Inches(w_q - 0.4), Inches(h_q - 0.4))
tf = tb_q3.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "3. Ongoing Security Management"
p.font.name = "Segoe UI"
p.font.size = Pt(20) 
p.font.bold = True
p.font.color.rgb = CYAN_ACCENT
p.space_after = Pt(8)
p_b1 = tf.add_paragraph()
p_b1.text = "• Header Hardening: Enable CSP & X-Frame-Options."
p_b1.font.name = "Segoe UI"
p_b1.font.size = Pt(15) 
p_b1.font.color.rgb = TEXT_WHITE
p_b2 = tf.add_paragraph()
p_b2.text = "• Auto updates: Configure automatic CMS updates."
p_b2.font.name = "Segoe UI"
p_b2.font.size = Pt(15) 
p_b2.font.color.rgb = TEXT_WHITE
p_b3 = tf.add_paragraph()
p_b3.text = "• Scanning: Run monthly scans for directory leaks."
p_b3.font.name = "Segoe UI"
p_b3.font.size = Pt(15) 
p_b3.font.color.rgb = TEXT_WHITE
draw_premium_card(slide8, x_r, y_b, w_q, h_q, accent_color=YELLOW_WARN)
tb_q4 = slide8.shapes.add_textbox(Inches(x_r + 0.2), Inches(y_b + 0.2), Inches(w_q - 0.4), Inches(h_q - 0.4))
tf = tb_q4.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "4. Lead Generation & Client Acquisition"
p.font.name = "Segoe UI"
p.font.size = Pt(20) 
p.font.bold = True
p.font.color.rgb = CYAN_ACCENT
p.space_after = Pt(8)
p_b1 = tf.add_paragraph()
p_b1.text = "• Lead Gen: Pitch audit reports to low-scoring sites."
p_b1.font.name = "Segoe UI"
p_b1.font.size = Pt(15) 
p_b1.font.color.rgb = TEXT_WHITE
p_b2 = tf.add_paragraph()
p_b2.text = "• Quick Wins: Offer free SSL & GBP verification setups."
p_b2.font.name = "Segoe UI"
p_b2.font.size = Pt(15) 
p_b2.font.color.rgb = TEXT_WHITE
p_b3 = tf.add_paragraph()
p_b3.text = "• Bundles: Package security audits for CA/clinics."
p_b3.font.name = "Segoe UI"
p_b3.font.size = Pt(15) 
p_b3.font.color.rgb = TEXT_WHITE
slide9 = prs.slides.add_slide(slide_layout)
apply_background(slide9)
create_title(slide9, "Lead Conversion & Outreach Playbook", "Commercial strategy for IT consulting & cybersecurity agencies")
col_w = 4.4
gap = 0.4
y_pos = 2.4
h_pos = 5.6
draw_premium_card(slide9, 1.0, y_pos, col_w, h_pos, accent_color=RED_RISK)
tb_c1 = slide9.shapes.add_textbox(Inches(1.2), Inches(y_pos + 0.3), Inches(col_w - 0.4), Inches(h_pos - 0.6))
tf = tb_c1.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "01  |  Lead Prioritization"
p.font.name = "Segoe UI"
p.font.size = Pt(26)
p.font.bold = True
p.font.color.rgb = RED_RISK
p.space_after = Pt(16)
p_body1_hdr = tf.add_paragraph()
p_body1_hdr.text = "Vulnerability Value"
p_body1_hdr.font.name = "Segoe UI"
p_body1_hdr.font.size = Pt(20)
p_body1_hdr.font.bold = True
p_body1_hdr.font.color.rgb = TEXT_WHITE
p_body1 = tf.add_paragraph()
p_body1.text = "• Target critical risks (score 0-2).\n• Focus on high-review businesses.\n• High reputation = high leverage."
p_body1.font.name = "Segoe UI"
p_body1.font.size = Pt(18)
p_body1.font.color.rgb = TEXT_MUTED
p_body1.space_before = Pt(4)
p_body1.space_after = Pt(16)
p_body2_hdr = tf.add_paragraph()
p_body2_hdr.text = "Sabotage Risks"
p_body2_hdr.font.name = "Segoe UI"
p_body2_hdr.font.size = Pt(20)
p_body2_hdr.font.bold = True
p_body2_hdr.font.color.rgb = TEXT_WHITE
p_body2 = tf.add_paragraph()
p_body2.text = "• Highlight unclaimed Maps listings.\n• Stress unauthorized edit risk."
p_body2.font.name = "Segoe UI"
p_body2.font.size = Pt(18)
p_body2.font.color.rgb = TEXT_MUTED
p_body2.space_before = Pt(4)
draw_premium_card(slide9, 1.0 + col_w + gap, y_pos, col_w, h_pos, accent_color=CYAN_ACCENT)
tb_c2 = slide9.shapes.add_textbox(Inches(1.2 + col_w + gap), Inches(y_pos + 0.3), Inches(col_w - 0.4), Inches(h_pos - 0.6))
tf = tb_c2.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "02  |  Conversion Hooks"
p.font.name = "Segoe UI"
p.font.size = Pt(26)
p.font.bold = True
p.font.color.rgb = CYAN_ACCENT
p.space_after = Pt(16)
p_body1_hdr = tf.add_paragraph()
p_body1_hdr.text = "Active Visual Warning"
p_body1_hdr.font.name = "Segoe UI"
p_body1_hdr.font.size = Pt(20)
p_body1_hdr.font.bold = True
p_body1_hdr.font.color.rgb = TEXT_WHITE
p_body1 = tf.add_paragraph()
p_body1.text = "• Show 'Not Secure' browser flags.\n• Demo exposed login directories.\n• Highlight SSL expiry warnings."
p_body1.font.name = "Segoe UI"
p_body1.font.size = Pt(18)
p_body1.font.color.rgb = TEXT_MUTED
p_body1.space_before = Pt(4)
p_body1.space_after = Pt(16)
p_body2_hdr = tf.add_paragraph()
p_body2_hdr.text = "Trust Builder Offer"
p_body2_hdr.font.name = "Segoe UI"
p_body2_hdr.font.size = Pt(20)
p_body2_hdr.font.bold = True
p_body2_hdr.font.color.rgb = TEXT_WHITE
p_body2 = tf.add_paragraph()
p_body2.text = "• Offer free Google Profile claim.\n• Charge low entry-fee for SSL setup."
p_body2.font.name = "Segoe UI"
p_body2.font.size = Pt(18)
p_body2.font.color.rgb = TEXT_MUTED
p_body2.space_before = Pt(4)
draw_premium_card(slide9, 1.0 + 2*(col_w + gap), y_pos, col_w, h_pos, accent_color=GREEN_SECURE)
tb_c3 = slide9.shapes.add_textbox(Inches(1.2 + 2*(col_w + gap)), Inches(y_pos + 0.3), Inches(col_w - 0.4), Inches(h_pos - 0.6))
tf = tb_c3.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "03  |  Upsell Strategy"
p.font.name = "Segoe UI"
p.font.size = Pt(26)
p.font.bold = True
p.font.color.rgb = GREEN_SECURE
p.space_after = Pt(16)
p_body1_hdr = tf.add_paragraph()
p_body1_hdr.text = "Compliance Focus"
p_body1_hdr.font.name = "Segoe UI"
p_body1_hdr.font.size = Pt(20)
p_body1_hdr.font.bold = True
p_body1_hdr.font.color.rgb = TEXT_WHITE
p_body1 = tf.add_paragraph()
p_body1.text = "• Pitch compliance audits to clinics.\n• Pitch secure bookings to hotels."
p_body1.font.name = "Segoe UI"
p_body1.font.size = Pt(18)
p_body1.font.color.rgb = TEXT_MUTED
p_body1.space_before = Pt(4)
p_body1.space_after = Pt(16)
p_body2_hdr = tf.add_paragraph()
p_body2_hdr.text = "Retainer Packages"
p_body2_hdr.font.name = "Segoe UI"
p_body2_hdr.font.size = Pt(20)
p_body2_hdr.font.bold = True
p_body2_hdr.font.color.rgb = TEXT_WHITE
p_body2 = tf.add_paragraph()
p_body2.text = "• Bundle weekly scanner audits.\n• Offer monthly backups & updates."
p_body2.font.name = "Segoe UI"
p_body2.font.size = Pt(18)
p_body2.font.color.rgb = TEXT_MUTED
p_body2.space_before = Pt(4)
prs.save("National_Digital_Risk_Audit_Presentation.pptx")
print("PowerPoint presentation National_Digital_Risk_Audit_Presentation.pptx generated successfully!")