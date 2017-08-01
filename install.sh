#!/bin/sh

if python3 -c "from pptx import Presentation" 2>/dev/null; then
  echo "python-pptx is already installed"
else
  echo "python-pptx is not installed"
  echo "python-pptx is installing now"
  pip3 install python-pptx
fi

chmod u+x image2ppt.py
cp image2ppt.py /usr/local/bin/image2ppt
echo "\n\033[92m  => image2ppt is installed globally. \033[0m\n"
