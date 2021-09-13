---
title: Элемент Hosts в файле манифеста
description: Указывает клиентское приложение Office, в котором будет активирована надстройка Office.
ms.date: 10/09/2018
ms.localizationpriority: medium
ms.openlocfilehash: 2684753fc32a295d7e177ef3bf668c194458128e
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151444"
---
# <a name="hosts-element"></a>Элемент Hosts

Указывает клиентское приложение Office, в котором будет активирована надстройка Office. Содержит коллекцию элементов **Host** и их параметров. 

При включении в узел [VersionOverrides](versionoverrides.md) этот элемент переопределяет элемент **Hosts** в родительской части манифеста. 

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Да   |  Описывает ведущее приложение и его параметры. |
