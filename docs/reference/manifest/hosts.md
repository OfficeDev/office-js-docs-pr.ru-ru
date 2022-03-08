---
title: Элемент Hosts в файле манифеста
description: Указывает Office клиентские приложения, Office надстройка активируется.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9ea6cc9745f47b6e9b1c9bb0232b744304078053
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63341074"
---
# <a name="hosts-element"></a>Элемент Hosts

Указывает Office клиентские приложения, Office надстройка активируется. Содержит коллекцию элементов **Host** и их параметров. 

## <a name="as-child-of-versionoverrides-element"></a>Как ребенок элемента VersionOverrides

Сведения в этом разделе применяются только *в том случае* , если элемент **Hosts** является ребенком [VersionOverrides](versionoverrides.md).

Этот элемент переопределяет **элемент Hosts** в базовом манифесте.

**Тип надстройки:** Области задач, Почта

**Допустимо только в этих схемах VersionOverrides**:

- Область задач 1.0
- Почта 1.0
- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Да   |  Описывает ведущее приложение и его параметры. |
