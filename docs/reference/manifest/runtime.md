---
title: Среда выполнения в файле манифеста
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: 68def44ba74733934198ac3b32fa1fe649156766
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111172"
---
# <a name="runtime-element"></a>Элемент среды выполнения

Эта функция доступна предварительная версия. Дочерний элемент [`<Runtimes>`](runtime.md) элемента. Этот элемент упрощает совместное использование глобальных данных и вызовов функций между пользовательскими функциями Excel и областью задач надстройки. 

## <a name="contained-in"></a>Содержится в

-[Сред выполнения](runtimes.md)

**Тип надстройки:** надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **время жизни = "Long"**  |  Да  | Всегда должен быть указан как длинное, если вы хотите, чтобы пользовательские функции Excel работали, когда область задач надстройки закрыта. |
|  **resid**  |  Да  | Если используется для пользовательских функций Excel, `resid` необходимо указать значение. `TaskPaneAndCustomFunction.Url` |

## <a name="see-also"></a>См. также

-[Полняющего](runtime.md)
