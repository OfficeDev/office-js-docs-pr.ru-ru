---
title: Среды выполнения в файле манифеста (Предварительная версия)
description: ''
ms.date: 02/21/2020
localization_priority: Normal
ms.openlocfilehash: 17e53b53d55ea9547cdfc5c4f89f8f4c3a7ab75e
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/26/2020
ms.locfileid: "42283879"
---
# <a name="runtimes-element-preview"></a>Элемент среды выполнения (Предварительная версия)

[!include[Running custom functions in browser runtime note](../../includes/excel-shared-runtime-preview-note.md)]

Указывает среду выполнения надстройки и позволяет использовать пользовательские функции, кнопки ленты и область задач для использования одной и той же среды выполнения JavaScript. Дочерний `<Host>` элемент элемента в файле манифеста. Дополнительные сведения можно найти в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](../../excel/configure-your-add-in-to-use-a-shared-runtime.md).

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
[Host](./host.md)

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Среда выполнения**     | Да |  Среда выполнения надстройки.

## <a name="see-also"></a>См. также

- [Среда выполнения](runtime.md)
