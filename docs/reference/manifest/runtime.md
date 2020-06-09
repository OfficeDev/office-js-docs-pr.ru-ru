---
title: Среда выполнения в файле манифеста
description: Элемент среды выполнения настраивает надстройку для использования общей среды выполнения JavaScript для различных компонентов, например ленты, области задач, настраиваемых функций.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: e81bd7222585bfa7d5f0f34fe5d9b32e4d45a71e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608106"
---
# <a name="runtime-element-preview"></a>Элемент среды выполнения (Предварительная версия)

Настраивает надстройку для использования общей среды выполнения JavaScript, чтобы различные компоненты запускались в одной среде выполнения. Дочерний [`<Runtimes>`](runtimes.md) элемент.

В Excel этот элемент позволяет использовать одну и ту же среду выполнения для ленты, области задач и пользовательских функций. Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

В Outlook этот элемент включает активацию надстройки на основе событий. Дополнительные сведения см. [в разделе Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md).

**Тип надстройки:** Область задач, почта

> [!IMPORTANT]
> **Excel**: общая среда выполнения в настоящее время доступна только в Excel для Windows.
>
> **Outlook**: Активация на основе событий в настоящее время находится [в предварительной версии](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) и доступна только в Outlook в Интернете. Дополнительные сведения см. [в статье Просмотр функции активации на основе событий](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).

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
|  **resid**  |  Да  | Указывает URL-адрес HTML-страницы для надстройки. `resid`Должен сопоставляться с `id` атрибутом `Url` элемента в `Resources` элементе. |
|  **время жизни**  |  Нет  | Значение по умолчанию для свойства `lifetime` `short` и не требуется указывать. В надстройках Outlook используется только `short` значение. Если вы хотите использовать общую среду выполнения в надстройке Excel, явно задайте для нее значение `long` . |

## <a name="see-also"></a>См. также

- [Runtimes](runtimes.md)
