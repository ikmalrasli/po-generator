from setuptools import setup, find_packages
import sys

setup(
    name="purchase-order-generator",
    version="1.0.0",
    description="Purchase Order Generator with AI-powered PDF processing",
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        "openpyxl>=3.1.0",
        "google-genai>=0.3.0", 
        "python-dotenv>=1.0.0",
        "num2words>=0.5.10",
        "tkcalendar>=1.6.1",
    ],
    python_requires=">=3.8",
    entry_points={
        'console_scripts': [
            'purchase-order-generator=main:main',
        ],
    },
)