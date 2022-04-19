import os, sys
import subprocess



key = ['L-VAD1','L-VAD2','L-VAD3','L-VAD4','L-VAD5','L-VAD7','L-VAD8','L-VAD9','L-VAD10','L-VAD11','L-VAD12','L-VAD13','L-VAD14',
'L-VAD15','L-VAD16','L-VAD17','L-VAD18','J-VAD07','J-VAD08','J-VAD09','J-VAD10','J-VAD11','J-VAD12','J-VAD13','J-VAD14','J-VAD15',
'J-VAD16','J-VAD17','J-VAD19','J-VAD20','O-VAD1','Tapering1','Sintering1','L-Lathe4','L-Lathe5','L-Lathe6','L-Lathe7','L-Lathe8',
'L-Lathe9','L-Lathe10','L-Lathe11','L-Lathe12','L-Lathe13','L-Furnace2','L-Furnace4','L-Furnace5','L-Furnace6','L-Furnace7','L-Furnace8','L-Furnace9','L-Furnace10','L-Furnace11','L-Furnace12','L-Furnace13','V-Furnace1','V-Furnace2','V-Furnace3','V-Furnace4',
'V-Lathe1','V-Lathe2','R-Furnace2','R-Furnace3','R-Furnace4','R-Furnace5','L-Drawing-1','L-Drawing-2','L-Drawing-3','L-Drawing-4',
'L-Drawing-5','L-Drawing-6','L-Drawing-7','L-Drawing-8','L-Drawing-9','L-Drawing-10','L-Drawing-11','L-Drawing-12','T-Drawing-13',
'T-Drawing-14','T-Drawing-15','T-Drawing-16','N-Drawing-17','N-Drawing-18','N-Drawing-19','N-Drawing-20','N-Drawing-21','N-Drawing-22','Rewinding1','Rewinding2','Rewinding4','Rewinding5','Rewinding6','Rewinding7','Rewinding8','Rewinding9','Rewinding10',
'Rewinding11','Rewinding12','Rewinding13','Rewinding14','Rewinding15','Rewinding16','Rewinding17','Rewinding25','Rewinding26',
'Rewinding27','Rewinding28','Rewinding29','Rewinding30','Rewinding31','Rewinding32','Rewinding33','Rewinding34','Rewinding35','Rewinding36','Rewinding38','Rewinding39','Rewinding40','Rewinding41','Rewinding42','Rewinding43','Rewinding44']



value=['10.12.14.6','10.12.14.7','10.12.14.8','10.12.14.9','10.12.14.10','10.12.14.12','10.12.14.13','10.12.14.14','10.12.14.15',
'10.12.14.16','10.12.14.17','10.12.14.18','10.12.14.19','10.12.14.20','10.12.14.22','10.12.14.24','10.12.14.26','10.12.14.63',
'10.12.14.64','10.12.14.65','10.12.14.66','10.12.14.67','10.12.14.68','10.12.14.69','10.12.14.70','10.12.14.71','10.12.14.72',
'10.12.14.73','10.12.14.74','10.12.14.76','10.12.14.78','10.12.14.79','10.12.14.80','10.12.14.93','10.12.14.96','10.12.14.97',
'10.12.14.98','10.12.14.94','10.12.14.99','10.12.14.105','10.12.14.95','10.12.14.101','10.12.14.102','10.12.14.36','10.12.14.38',
'10.12.14.39','10.12.14.40','10.12.14.41','10.12.14.42','10.12.14.43','10.12.14.44','10.12.14.45','10.12.14.46','10.12.14.48',
'10.12.14.84','10.12.14.85','10.12.14.86','10.12.14.87','10.12.14.103','10.12.14.104','10.12.14.53','10.12.14.57','10.12.14.54',
'10.12.14.59','10.12.14.108','10.12.14.109','10.12.14.110','10.12.14.111','10.12.14.112','10.12.14.113','10.12.14.114',
'10.12.14.115','10.12.14.116','10.12.14.117','10.12.14.118','10.12.14.119','10.12.14.127','10.12.14.128','10.12.14.129','10.12.14.130','10.12.14.120','10.12.14.121','10.12.14.122','10.12.14.123','10.12.14.124','10.12.14.125','10.12.14.135','10.12.14.136',
'10.12.14.137','10.12.14.138','10.12.14.139','10.12.14.140','10.12.14.141','10.12.14.142','10.12.14.143','10.12.14.144',
'10.12.14.145','10.12.14.146','10.12.14.147','10.12.14.148','10.12.14.149','10.12.14.150','10.12.14.175','10.12.14.175','10.12.14.151','10.12.14.152','10.12.14.153','10.12.14.154','10.12.14.155','10.12.14.156','10.12.14.161','10.12.14.162','10.12.14.163',
'10.12.14.164','10.12.14.176','10.12.14.177','10.12.14.178','10.12.14.157','10.12.14.158','10.12.14.159','10.12.14.160']


dict = { name:value for name, value in zip(key,value)}

def check(ip,file):
  param = '-n'
  command = ['ping', param, '1', ip]
  process = subprocess.Popen(command, stdout=subprocess.PIPE);
  for line in iter(process.stdout.readline, b''):
      sys.stdout.write(str(line));
      file.write(str(line));


with open('C:/Users/user/OneDrive - 대한광통신/바탕 화면/크롤링/t.log', 'w') as file:
  for k,v in dict.items():
    print(check(v,file));


  
  # param = '-n'
  # command = ['ping', param, '1', ip]

