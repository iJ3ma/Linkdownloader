from setuptools import setup, find_packages

setup(
    name='excel-link-downloader',
    version='0.1.0',
    author='Your Name',
    author_email='your.email@example.com',
    description='A tool for downloading links from Excel files and websites.',
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/yourusername/excel-link-downloader',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    install_requires=[
        'openpyxl',
        'requests',
        'beautifulsoup4',
    ],
    entry_points={
        'console_scripts': [
            'excel-link-downloader=excel_downloader:main',  # Adjust if main function is defined
        ],
    },
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',
)