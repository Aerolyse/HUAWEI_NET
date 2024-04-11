# Documentation pour `HUAWEI_NET.py`

Ce document fournit une vue d'ensemble et une explication détaillée du fichier `HUAWEI_NET.py`, utilisé pour interagir avec les équipements réseaux de Huawei via des requêtes API et pour gérer les configurations réseau.

## Index

1. [Introduction](#introduction)
2. [Fonctions Principales](#fonctions-principales)
   - [get_token](#get_token)
   - [get_device](#get_device)
   - [Device_Sorter](#device_sorter)
   - [export_xlsx](#export_xlsx)
   - [update_ethernet_request](#update_ethernet_request)
   - [update_ethtrunk_request](#update_ethtrunk_request)
   - [get_interfaces](#get_interfaces)
   - [dict_comparator](#dict_comparator)
   - [type_converter](#type_converter)
   - [ethernet_value_comparator](#ethernet_value_comparator)
   - [ethtrunk_value_comparator](#ethtrunk_value_comparator)
   - [Ethernet_Request](#ethernet_request)
   - [EthTrunk_Request](#ethtrunk_request)
   - [Import](#import)
3. [Utilisation](#utilisation)

## Introduction

Le script `HUAWEI_NET.py` est conçu pour automatiser plusieurs tâches de gestion des équipements réseau Huawei à travers des interactions avec l'API de Huawei. Il permet de récupérer des informations sur les dispositifs, de trier ces dispositifs selon certaines spécifications, d'exporter ces informations sous format Excel, et de mettre à jour la configuration des ports Ethernet.

## Fonctions Principales

### `get_token`

**Description**:
Obtient un token d'authentification de l'API Huawei.

**Paramètres**:
- Aucun

**Retourne**:
- `token_id_value`: Token utilisé pour les requêtes authentifiées.

**Exemple d'utilisation**:

```python
token = get_token()
```

### `get_device`

**Description**:
Récupère les dispositifs depuis l'API Huawei en utilisant un token d'authentification.

**Paramètres**:
- `token`: Token d'authentification obtenu via `get_token`.

**Retourne**:
- `device_dict`: Dictionnaire contenant les données des dispositifs.

**Exemple d'utilisation**:

```python
devices = get_device(token)
```

### `Device_Sorter`

**Description**:
Trie les dispositifs par type et nom de site.

**Paramètres**:
- `site_Name_args`: Nom du site pour lequel trier les dispositifs.
- `device_dict`: Dictionnaire des dispositifs.
- `device_Type_args`: Type de dispositif à filtrer (par défaut "LSW").

**Retourne**:
- `devices_List`: Liste des dispositifs triés.

**Exemple d'utilisation**:

```python
sorted_devices = Device_Sorter("Site1", devices)
```

### `export_xlsx`

**Description**:
Exporte une liste de dispositifs triés dans un fichier Excel.

**Paramètres**:
- `token`: Token d'authentification.
- `site_devices_List`: Liste des dispositifs à exporter.
- `custom_Name_Path`: Chemin du fichier de sortie.

**Retourne**:
- Aucun

**Exemple d'utilisation**:

```python
export_xlsx(token, sorted_devices, "output.xlsx")
```

### `update_ethernet_request`

**Description**:
Met à jour la configuration Ethernet d'un dispositif spécifique.

**Paramètres**:
- `token`: Token d'authentification.
- `body_dict`: Dictionnaire de la configuration à appliquer.
- `Id`: Identifiant du dispositif.
- `Device_name`: Nom du dispositif.
- `ethernet_interface_Name`: Nom de l'interface Ethernet.

**Retourne**:
- Aucun

**Exemple d'utilisation**:

```python
update_ethernet_request(token, config, device_id, "Device1", "eth0")
```

### `update_ethtrunk_request`

**Description**:
Met à jour la configuration Eth-Trunk d'un dispositif spécifique.

**Paramètres**:
- `token`: Token d'authentification.
- `body`: Dictionnaire de la configuration à appliquer.
- `Id`: Identifiant du dispositif.
- `Device_name`: Nom du dispositif.
- `ethtrunk_name`: Nom de l'Eth-Trunk.

**Retourne**:
- Aucun

**Exemple d'utilisation**:

```python
update_ethtrunk_request(token, trunk_config, device_id, "Device1", "trunk1")
```

### `get_interfaces`

**Description**:
Obtient les interfaces réseau d'un dispositif.

**Paramètres**:
- `token`: Token d'authentification.
- `Id`: Identifiant du dispositif.
- `interface_Type`: Type d'interface ("Ethernet" ou "Eth-Trunk").

**Retourne**:
- Liste des interfaces.

**Exemple d'utilisation**:

```python
interfaces = get_interfaces(token, device_id, "Ethernet")
```

### `dict_comparator`

**Description**:
Compare deux dictionnaires pour s'assurer que le dictionnaire à envoyer à l'API contient uniquement les clés attendues.

**Paramètres**:
- `compared_dict`: Dictionnaire à comparer.
- `comparator_dict`: Dictionnaire de comparaison.
- `interface_Type`: Type d'interface.

**Retourne**:
- `compared_dict`: Dictionnaire nettoyé.

**Exemple d'utilisation**:

```python
clean_dict = dict_comparator(new_config, existing_config, "Ethernet")
```

### `type_converter`

**Description**:
Convertit les types de données dans un dictionnaire pour correspondre aux types attendus par l'API.

**Paramètres**:
- `interface_comparator`: Dictionnaire comparateur.
- `interface_compared`: Dictionnaire à comparer.
- `interface_Type`: Type d'interface.

**Retourne**:
- Dictionnaire avec les types de données ajustés.

**Exemple d'utilisation**:

```python
converted_dict = type_converter(comparator, to_convert, "Ethernet")
```

### `ethernet_value_comparator`

**Description**:
Compare les valeurs de deux dictionnaires pour les interfaces Ethernet.

**Paramètres**:
- `interfaces_compared`: Dictionnaire des interfaces comparées.
- `interfaces_comparator`: Dictionnaire comparateur.
- `interface_Type`: Type d'interface.

**Retourne**:
- Booléen indiquant s'il existe des différences.

**Exemple d'utilisation**:

```python
has_diff = ethernet_value_comparator(current_config, new_config, "Ethernet")
```

### `ethtrunk_value_comparator`

**Description**:
Compare les valeurs de deux dictionnaires pour les interfaces Eth-Trunk.

**Paramètres**:
- `ethtrunk_compared`: Dictionnaire des interfaces comparées.
- `ethtrunk_comparator`: Dictionnaire comparateur.
- `interface_Type`: Type d'interface.

**Retourne**:
- Booléen indiquant s'il existe des différences.

**Exemple d'utilisation**:

```python
has_diff = ethtrunk_value_comparator(current_trunk_config, new_trunk_config, "Eth-Trunk")
```

### `Ethernet_Request`

**Description**:
Traite une demande de configuration Ethernet et met à jour si nécessaire.

**Paramètres**:
- `token`: Token d'authentification.
- `stackId`: Identifiant de la pile.
- `device_name`: Nom du dispositif.
- `deviceId`: Identifiant du dispositif.
- `device_interface_list`: Liste des interfaces du dispositif.
- `request_interface_dict`: Dictionnaire de la demande d'interface.
- `interface_Type`: Type d'interface.
- `ethernet_interface_Name`: Nom de l'interface Ethernet.
- `already_seen_devices_list`: Liste des dispositifs déjà traités.

**Retourne**:
- Liste contenant les informations sur les modifications détectées et mises à jour.

**Exemple d'utilisation**:

```python
result = Ethernet_Request(token, stack_id, "Device1", device_id, interfaces, config, "Ethernet", "eth0", seen_devices)
```

### `EthTrunk_Request`

**Description**:
Traite une demande de configuration Eth-Trunk et met à jour si nécessaire.

**Paramètres**:
- `token`: Token d'authentification.
- `stackId`: Identifiant de la pile.
- `device_name`: Nom du dispositif.
- `deviceId`: Identifiant du dispositif.
- `device_ethtrunk_list`: Liste des Eth-Trunks du dispositif.
- `request_ethtrunk_dict`: Dictionnaire de la demande d'Eth-Trunk.
- `interface_Type`: Type d'interface.
- `ethtrunk_interface_Name`: Nom de l'Eth-Trunk.
- `already_seen_devices_list`: Liste des dispositifs déjà traités.

**Retourne**:
- Liste contenant les informations sur les modifications détectées et mises à jour.

**Exemple d'utilisation**:

```python
result = EthTrunk_Request(token, stack_id, "Device1", device_id, trunks, trunk_config, "Eth-Trunk", "trunk1", seen_devices)
```

### `Import`

**Description**:
Importe les configurations d'un fichier Excel et applique les modifications nécessaires.

**Paramètres**:
- `token`: Token d'authentification.
- `filename_or_path`: Chemin du fichier Excel à importer.

**Retourne**:
- Aucun

**Exemple d'utilisation**:

```python
Import(token, "config.xlsx")
```

## Utilisation

Pour utiliser ce script, vous devez d'abord obtenir un token avec la fonction `get_token()`, puis vous pouvez appeler les autres fonctions en passant ce token comme argument. Le script peut être exécuté directement avec des arguments spécifiques pour exporter ou importer des configurations via des fichiers Excel.

**Exemple de commande**:

```bash
python HUAWEI_NET.py -s Site1 -f output.xlsx
```

Cette commande exporte les configurations du site `Site1` dans le fichier `output.xlsx`.
