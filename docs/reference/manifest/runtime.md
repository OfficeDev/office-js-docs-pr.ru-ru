---
title: Среда выполнения в файле манифеста (Предварительная версия)
description: Элемент среды выполнения настраивает надстройку для использования общей среды выполнения JavaScript для ленты, области задач и пользовательских функций.
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 6237f64fec47ed22b0105bf74c8eb7e2b7c38afe
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717931"
---
# <a name="runtime-element-preview"></a>Элемент среды выполнения (Предварительная версия)

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

Дочерний элемент [`<Runtimes>`](runtimes.md) элемента. Этот элемент настраивает надстройку, чтобы использовать общую среду выполнения JavaScript, чтобы Ваша лента, область задач и пользовательские функции выполнялись в одной среде выполнения. Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

**Тип надстройки:** надстройки области задач

> [!IMPORTANT]
> Общедоступная среда выполнения в настоящее время находится в режиме предварительной версии и доступна только в Excel для Windows. Для ознакомления с предварительными возможностями необходимо присоединиться к [программе предварительной оценки Office](https://insider.office.com/).

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
|  **время жизни = "Long"**  |  Да  | Всегда следует использовать `long` , если вы хотите использовать общую среду выполнения для надстройки Excel. |
|  **resid**  |  Да  | Указывает URL-адрес HTML-страницы для надстройки. `resid` Должен сопоставляться с `id` атрибутом `Url` элемента в `Resources` элементе. |

## <a name="see-also"></a>См. также

- [Runtimes](runtimes.md)
