from setuptools import setup, find_packages

setup(
    name="smallworld-automation",
    version="1.0.0",
    author="Enterprise",
    description="Smallworld GIS Automation Platform",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    python_requires=">=3.9",
)
