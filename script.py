import os
os.chdir('/home/ubuntu/Zeiss_RFQ_tracker')
os.system("git reset --hard")
os.system("git pull")
os.system ('sudo chown :www-data ~/Zeiss_RFQ_tracker/')
os.system('sudo chmod 777 -R  ~/Zeiss_RFQ_tracker/')
os.system ('sudo systemctl restart apache2')
os.system ('cat /etc/apache2/envvars')
