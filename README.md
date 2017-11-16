# Создание-Implementation-plans

Данный скрипт генерирует планы используя шаблон word
В скрипте присутсвуют дебажные принты, так как информация разрознена, собиралась разными людьми и не приведена к общиму знаменателю.
Скриптсоздан для упрощении работы. Сохранения личного времени. Иначе пришлось бы создавать 111 файлов полностью руками.

# Устанока

```bash
pip install -r requirements.txt # alternatively try pip3
```

# Описание функций

generate_template(tpl, site, site_information)
создает словарь переменных для шаблона,на вход принимает название сайта, и вводную иноформацию по сайту.

get_sites_list()
возвращает спсиок названий сайтов

genetate_implementation_design(context,  tpl)
создает implementation_design используя словарь переменных и шаблон.

get_site_address(site)
возвращает адрес сайта и его координаты. Данные берутся из xls файла

get_site_device(site)
возвращает текушие устройсво и устройство на которое оно будет заменено. Данные берутся из xls файла

get_sites_information()
возвращает данные по сайту. Данные берутся из xls файла

get_software(site, site_information)

get_topology(ssh)
отсылает команду show interfaces  status и затем парсит ее

get_topology_from_file(site)
то же самое что и выше, но из файла

get_configuration(old_hostname, ssh)
собирает текущую конфигурацию с циски. 
и целовй конифг хуавей

connect_to_device(ip)
непосредственно подключение к устройсту


send_show_command(ssh_connection, command):
    отправление уоманды 


 disconnect(ssh_connection):
  дисконект



def get_vlan_list(ssh):
создание списка влан из вывода команды show vlan brief


def get_vlan_int_list(ssh):
создание списка влан из вывода команды show ip interface brief | include lan
    
def parse_vlan_int_conf(ssh, vlan_id):
 проверяет наличие IP, VRF , Xconect на интерфейсе


def generate_vlan_table_content(ssh):
создает данне для таблицы vlan




# Пример запуска скрипата


```bash
IPBB_Mrgashat
host: Mrgashat
 current d: Cisco ASR-901-6CZ-FT-D
 target d: Huawei ATN 910B)
hostname am0303ro2
Loopback: 10.6.0.80  /32
connected to 10.6.0.80
gather  interfaces  status  
collect vlan info
collect vlan int list
{'1002', '1004', '1005', '1003'}
gather configuration from:
IPBB-vill.Armavir
host: Armavir
 current d: Cisco 7606
 target d: Huawei CX600-X8)
hostname am0303ro3
Loopback: 10.6.0.81  /32
connected to 10.6.0.81
gather  interfaces  status  
collect vlan info
collect vlan int list
{'1002', '1004', '1005', '1003'}
gather configuration from:

```
