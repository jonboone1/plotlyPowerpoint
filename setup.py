import setuptools 

REQUIREMENTS = [i.strip() for i in open("requirements.txt").readlines()]

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="plotlyPowerpoint",
    version="1.2.6",
    author="Jon Boone",
    author_email="jonboone1@gmail.com",
    description="A library using Plotly and Powerpoint to easily generate slides with plotly charts in them",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/jonboone1/plotlyPowerpoint",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent"
    ],
    python_requires=">=3.6",
    install_requires=REQUIREMENTS
)