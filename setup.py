from setuptools import setup, find_packages

setup(
    name="proposta_solar",
    version="1.0.0",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=[
        "pandas>=2.0.0",
        "matplotlib>=3.10.0",
        "python-pptx>=0.6.21",
        "openpyxl>=3.1.0",
        "numpy>=1.24.0"
    ],
    python_requires=">=3.8",
    entry_points={
        'console_scripts': [
            'gerar-proposta=proposta_solar.cli:main',
        ],
    },
    author="José Oliveira",
    description="Gerador de propostas para energia solar",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
) 