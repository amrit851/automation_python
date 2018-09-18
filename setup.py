from distutils.core import setup
import py2exe, sys, os
 
py2exe_options = dict( 
						            # excludes= [],  # Exclude standard library
                        excludes= ['_ssl','email', 'urllib', 'doctest','pdb','unittest','inspect'],
                        dll_excludes = ['msvcr71.dll'],  # Exclude msvcr71
                        compressed = True,  # Compress library.zip
                      )

# sound_files = ['sounds/bye.wav', 'sounds/welcome.wav','sounds/short_while.wav']
# help_files = ['help/functions-help.htm', 'help/main-help.htm']
other_files = ['cred.py']
image_files = []
folder = "C:\\Users\\1023495\\Documents\\python files\\salesforce_scrape_script_schange_automation\\img"
for file in os.listdir(folder):
	if os.path.isfile(folder+"\\"+file):
		image_files.append('img\\'+file)

# image_files = ['images/logo.ico', 'images/logo.gif']
# print (len(image_files))

setup(name='RICHI',
      version='1.0.0',
      description='Application which automates salesforce dataclear case request',
      author='Amritanshu',
      author_email='....@gmail.com',
      data_files=[('', other_files), ('img', image_files)],
      console=['salescrape_script_change.py'],
      options={'py2exe': py2exe_options},
      )