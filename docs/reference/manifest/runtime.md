---
title: Среда выполнения в файле манифеста
description: Элемент среды выполнения настраивает надстройку для использования общей среды выполнения JavaScript для ленты, области задач и пользовательских функций.
ms.date: 05/11/2020
localization_priority: Normal
ms.openlocfilehash: c5c7356f9985ca7b5972068629b0587f8916348e
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/13/2020
ms.locfileid: "44217762"
---
# <a name="runtime-element"></a>Элемент среды выполнения

Дочерний элемент [`<Runtimes>`](runtimes.md) элемента. Этот элемент настраивает надстройку, чтобы использовать общую среду выполнения JavaScript, чтобы Ваша лента, область задач и пользовательские функции выполнялись в одной среде выполнения. Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

**Тип надстройки:** надстройки области задач

## <a name="syntax"></a>Синтаксис

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Содержится в

- [Runtimes](runtimes.md)

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **время жизни = "Long"**  |  Да  | Всегда следует `long` использовать, если вы хотите использовать общую среду выполнения для надстройки Excel. |
|  **resid**  |  Да  | Указывает URL-адрес HTML-страницы для надстройки. `resid`Должен сопоставляться с `id` атрибутом `Url` элемента в `Resources` элементе. |

## <a name="see-also"></a>См. также

- [Runtimes](runtimes.md)
