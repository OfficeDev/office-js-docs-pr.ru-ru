---
title: Элемент Hosts в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 606073977366e37ecc4419f468f01bfb25647a7d
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452027"
---
# <a name="hosts-element"></a>Элемент Hosts

Указывает клиентское приложение Office, в котором будет активирована надстройка Office. Содержит коллекцию элементов **Host** и их параметров. 

При включении в узел [VersionOverrides](versionoverrides.md) этот элемент переопределяет элемент **Hosts** в родительской части манифеста. 

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Да   |  Описывает ведущее приложение и его параметры. |
