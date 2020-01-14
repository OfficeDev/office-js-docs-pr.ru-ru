---
title: Среды выполнения в файле манифеста
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: ec2b85a92325eb4e36c61f731369ec54d44ef169
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111179"
---
# <a name="runtimes-element"></a>Элемент среды выполнения

Эта функция доступна предварительная версия. Определяет среду выполнения надстройки и позволяет использовать пользовательские функции и область задач для совместного использования глобальных данных и выполнения вызовов функций друг на друга. Должен следовать `<Host>` элементу в файле манифеста.

**Тип надстройки:** надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Среда выполнения**     | Да |  Среда выполнения надстройки, часто используемая с пользовательскими функциями Excel.

## <a name="see-also"></a>См. также

-[Сред выполнения](runtimes.md)
