---
title: Элемент Hosts в файле манифеста
description: Указывает клиентское приложение Office, в котором будет активирована надстройка Office.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 037ac2b5fedbfb1b59b7523382574942fe59a00a
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938053"
---
# <a name="hosts-element"></a>Элемент Hosts

Указывает клиентское приложение Office, в котором будет активирована надстройка Office. Содержит коллекцию элементов **Host** и их параметров. 

При включении в узел [VersionOverrides](versionoverrides.md) этот элемент переопределяет элемент **Hosts** в родительской части манифеста. 

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Да   |  Описывает ведущее приложение и его параметры. |
