---
title: Среды выполнения в файле манифеста
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 6682887935ee6894b5a311ad519408067452bb23
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554008"
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

- [Среда выполнения](runtime.md)
