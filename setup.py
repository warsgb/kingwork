#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""KingWork 安装脚本"""
from setuptools import setup, find_packages

setup(
    name="kingwork",
    version="1.0.0",
    description="KingWork - 工作智能管理技能",
    packages=find_packages(),
    python_requires=">=3.6",
    install_requires=[
        "requests>=2.28.0",
        "pyyaml>=6.0",
        "python-dateutil>=2.8.0",
    ],
    entry_points={
        "console_scripts": [
            "kw=skills.kingrecord.run:main",
        ],
    },
)
