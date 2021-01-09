---
title: Runtime в файле манифеста
description: Элемент runtime настраивает надстройку на использование общей компоненты javaScript для различных компонентов, например ленты, области задач, пользовательских функций.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 3cabfacc665ccf6c0e4e796cb0e1fbc70c770ee3
ms.sourcegitcommit: 545888b08f57bb1babb05ccfd83b2b3286bdad5c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2021
ms.locfileid: "49789186"
---
# <a name="runtime-element-preview"></a>Элемент runtime (предварительная версия)

Настраивает надстройку для использования общей времени работы JavaScript, чтобы все компоненты запускались в одной среде. Child of the [`<Runtimes>`](runtimes.md) element.

В Excel этот элемент позволяет ленте, области задач и пользовательским функциям использовать ту же времени работы. Дополнительные сведения см. в настройках надстройки Excel для использования общей времени [работы JavaScript.](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)

В Outlook этот элемент включает активацию надстройки на основе событий. Дополнительные сведения см. в настройке [надстройки Outlook для активации на основе событий.](../../outlook/autolaunch.md)

**Тип надстройки:** Области задач, почта

> [!IMPORTANT]
> **Outlook**: активация на основе событий в настоящее время находится в [предварительной](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) версии и доступна только в Outlook в Интернете. Дополнительные сведения см. в [предварительном просмотре функции активации на основе событий.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)

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
|  **resid**  |  Да  | Указывает URL-адрес HTML-страницы для надстройки. Он может иметь не более 32 символов и должен соответствовать `resid` `id` атрибуту `Url` элемента в `Resources` элементе. |
|  **lifetime**  |  Нет  | Значение по умолчанию : и не `lifetime` `short` требуется быть заданным. Надстройки Outlook используют только `short` значение. Если вы хотите использовать общую time runtime в надстройки Excel, явно установите значение `long` . |

## <a name="see-also"></a>См. также

- [Runtimes](runtimes.md)
