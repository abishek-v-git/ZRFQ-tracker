import os
os.chdir('/home/ubuntu/ZRFQ-tracker')
os.system("git reset --hard")
os.system("git pull")
os.system ('sudo chown :www-data ~/ZRFQ-tracker/')
os.system('sudo chmod 777 -R  ~/ZRFQ-tracker/')
os.system ('sudo systemctl restart apache2')
os.system ('cat /etc/apache2/envvars')
