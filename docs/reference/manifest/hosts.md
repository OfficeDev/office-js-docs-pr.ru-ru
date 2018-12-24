---
title: Элемент Hosts в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 59010c0f6c0d14d8721856f81def11540db28704
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433413"
---
# <a name="hosts-element"></a>Элемент Hosts

Указывает клиентское приложение Office, в котором будет активирована надстройка Office. Содержит коллекцию элементов **Host** и их параметров. 

При включении в узел [VersionOverrides](versionoverrides.md) этот элемент переопределяет элемент **Hosts** в родительской части манифеста. 

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Да   |  Описывает ведущее приложение и его параметры. |
