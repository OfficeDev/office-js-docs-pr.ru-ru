---
title: Элемент Hosts в файле манифеста
description: Указывает клиентское приложение Office, в котором будет активирована надстройка Office.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cd4e0eecce610b10fdc9dafcde7b807fde425b14
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718106"
---
# <a name="hosts-element"></a>Элемент Hosts

Указывает клиентское приложение Office, в котором будет активирована надстройка Office. Содержит коллекцию элементов **Host** и их параметров. 

При включении в узел [VersionOverrides](versionoverrides.md) этот элемент переопределяет элемент **Hosts** в родительской части манифеста. 

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Да   |  Описывает ведущее приложение и его параметры. |
