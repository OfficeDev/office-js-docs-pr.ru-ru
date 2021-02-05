---
title: Runtimes in the manifest file
description: Элемент Runtimes указывает времени работы надстройки.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 74bb2b432f46d5876601052003e20ff843e13b06
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104828"
---
# <a name="runtimes-element"></a>Элемент Runtimes

Указывает времени работы надстройки. Child of the [`<Host>`](host.md) element.

> [!NOTE]
> При запуске в Office для Windows надстройка использует браузер Internet Explorer 11.

В Excel этот элемент позволяет ленте, области задач и пользовательским функциям использовать ту же времени работы. Дополнительные сведения см. в настройках надстройки Excel для использования общей времени [работы JavaScript.](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)

В Outlook этот элемент включает активацию надстройки на основе событий. Дополнительные сведения см. в настройке [надстройки Outlook для активации на основе событий.](../../outlook/autolaunch.md)

**Тип надстройки:** Области задач, почта

> [!IMPORTANT]
> **Outlook**: функция активации на [](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) основе событий в настоящее время находится в предварительной версии и доступна только в Outlook в Интернете и Windows. Дополнительные сведения см. в [предварительном просмотре функции активации на основе событий.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)

## <a name="syntax"></a>Синтаксис

```XML
<Runtimes>
    <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a>Содержится в

[Host](host.md)

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
| [Runtime](runtime.md) | Да |  Время работы надстройки. |

## <a name="see-also"></a>См. также

- [Runtime](runtime.md)
