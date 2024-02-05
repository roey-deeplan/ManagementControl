#!/bin/bash

# Run the build
node updateVersion.service.js && npm run clean && npm run bundle && npm run package-solution && npm run open-explorer

# Check if the build was successful
if [ $? -eq 0 ]; then
    echo "Build successful"
else
    echo "Build failed, reverting version number"
    node decreaseVersion.service.js
fi
