# Lucee Extension
This repository holds the source code for an Extension that can be installed into a regular Lucee Installation.
It provides Apache POI functionlaity.

## Build the Extension
This build process uses "ant" to build the Extension, so make sure you have installed "ant" on your system.
then go to the root folder of the project on your command line and call "ant", that's it.
After sucessfully done, the built extension (.lex file) will be available in the folder "dist".

## Install the Extension
Copy the file *.lex from to the dist folder to "/lucee-server/deploy" of your Lucee Installation, when Lucee runs it will within a minute pick up the extension and install it on the fly (no restart necessary). The extension can also be installed with help of an extension provider at startup with help of enviroment variables. Please ask Michael Offner when you need more details on this.
