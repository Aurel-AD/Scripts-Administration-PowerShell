# 🖥️ WhatIsMyIP_Address

## 📝 Description

Outil de diagnostic Helpdesk permettant d'afficher rapidement les principales informations d'un poste Windows. Il facilite les opérations de support, d'inventaire et d'identification matérielle lors des interventions de niveau 1 et 2.

## ⚙️ Prérequis

- Windows 10 ou Windows 11
- PowerShell 5.1 ou supérieur
- Aucun droit administrateur requis pour l'utilisation standard

## 🚀 Utilisation

```powershell
.\WhatIsMyIP_Address.ps1
```

## 📋 Informations collectées

### 🌐 Réseau
- Adresse IPv4 principale
- Adresse IPv4 secondaire (si disponible)

### 💻 Système
- Nom de l'ordinateur
- Version de Windows
- Architecture du système (32/64 bits)
- Date d'installation du système
- Temps de fonctionnement (uptime)

### 🔧 Matériel
- Fabricant et modèle du poste
- Tag d'inventaire / numéro de série
- Version du BIOS
- Quantité de mémoire RAM installée
- Modèle du processeur
- Informations du disque principal (taille, modèle et numéro de série)

## ✅ Cas d'usage

- Diagnostic rapide lors d'un appel utilisateur
- Inventaire matériel d'un poste
- Identification d'une machine à distance
- Vérification de la configuration système
- Support Helpdesk N1 / N2

## 📷 Exemple de sortie

```text
==============================================================
My IP Address1 is : 192.168.1.10
My IP Address2 is : n/a

My ComputerName is : PC-LENS01
==============================================================

Hardware    make|model : Dell Inc. | Latitude 5540
           tag|bios : ABC123 | 1.20.0
           memory : 16GB
           cpu : Intel Core i5-1345U
           disk1 : 512GB - NVMe SSD - s/n: XXXXXXXX
--------------------------------------------------------------
Software

           OS : Microsoft Windows 11 Enterprise, Version 24H2, 64-bit
   install date : 2025/09/15 08:30
        uptime : 5 day(s), 3 hour(s), 22 minute(s), 10 second(s).
==============================================================


