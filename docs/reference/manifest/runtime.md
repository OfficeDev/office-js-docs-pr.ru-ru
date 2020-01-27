---
title: Среда выполнения в файле манифеста
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 8fbad8276b3e1d64a6c443cf57d498597d729282
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554001"
---
# <a name="runtime-element"></a>Элемент среды выполнения

Эта функция доступна предварительная версия. Дочерний элемент [`<Runtimes>`](runtimes.md) элемента. Этот элемент упрощает совместное использование глобальных данных и вызовов функций между пользовательскими функциями Excel и областью задач надстройки.

**Тип надстройки:** надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Содержится в

- [Runtimes](runtimes.md)

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **время жизни = "Long"**  |  Да  | Всегда должен быть указан как длинное, если вы хотите, чтобы пользовательские функции Excel работали, когда область задач надстройки закрыта. |
|  **resid**  |  Да  | Если используется для пользовательских функций Excel, `resid` необходимо указать значение. `TaskPaneAndCustomFunction.Url` |

## <a name="see-also"></a>См. также

- [Runtimes](runtimes.md)
