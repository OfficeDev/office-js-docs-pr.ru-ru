---
title: Среда выполнения в файле манифеста (Предварительная версия)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 26702896604f9ecf4c69296e5110efe5cdf4218b
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283886"
---
# <a name="runtime-element-preview"></a>Элемент среды выполнения (Предварительная версия)

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

Дочерний элемент [`<Runtimes>`](runtimes.md) элемента. Этот элемент настраивает надстройку, чтобы использовать общую среду выполнения JavaScript, чтобы Ваша лента, область задач и пользовательские функции выполнялись в одной среде выполнения. Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

**Тип надстройки:** надстройки области задач

> [!IMPORTANT]
В настоящее время общедоступная среда выполнения <<<<<<< для ГОЛОВного общего доступа доступна только в Excel для Windows. Для ознакомления с предварительными возможностями необходимо присоединиться к [программе предварительной оценки Office](https://insider.office.com/).

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
