---
title: Среды выполнения в файле манифеста
description: Элемент Runtimes указывает среду выполнения надстройки.
ms.date: 06/01/2020
localization_priority: Normal
ms.openlocfilehash: 082491befc6b9dbdc474b0e40f9defd90a4ef75f
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159362"
---
# <a name="runtimes-element"></a>Элемент среды выполнения

Задает среду выполнения надстройки. Дочерний [`<Host>`](host.md) элемент.

> [!NOTE]
> При работе в Office в Windows надстройка использует браузер Internet Explorer 11.

В Excel этот элемент позволяет использовать одну и ту же среду выполнения для ленты, области задач и пользовательских функций. Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

В Outlook этот элемент включает активацию надстройки на основе событий. Дополнительные сведения см. [в разделе Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md).

**Тип надстройки:** Область задач, почта

> [!IMPORTANT]
> **Outlook**: функция активации на основе событий в настоящее время находится [в предварительной версии](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) и доступна только в Outlook в Интернете. Дополнительные сведения см. [в статье Просмотр функции активации на основе событий](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).

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
| [Время выполнения](runtime.md) | Да |  Среда выполнения надстройки. |

## <a name="see-also"></a>См. также

- [Время выполнения](runtime.md)
