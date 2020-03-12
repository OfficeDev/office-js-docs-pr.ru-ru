---
title: Элемент Enabled в файле манифеста
description: Сведения о том, как указать, что команда надстройки отключена при запуске надстройки.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 4c2c013c8e55966ba2678755536ce04ae3014ed0
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596902"
---
# <a name="enabled-element"></a><span data-ttu-id="27d0c-103">Элемент Enabled</span><span class="sxs-lookup"><span data-stu-id="27d0c-103">Enabled element</span></span>

<span data-ttu-id="27d0c-104">Указывает, включен ли элемент управления ["Кнопка" или "](control.md#button-control) [меню](control.md#menu-dropdown-button-controls) " при запуске надстройки.</span><span class="sxs-lookup"><span data-stu-id="27d0c-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="27d0c-105">Элемент **Enabled** является дочерним элементом [элемента управления](control.md).</span><span class="sxs-lookup"><span data-stu-id="27d0c-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="27d0c-106">Если он не указан, используется `true`значение по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="27d0c-106">If it is omitted, the default is `true`.</span></span>

<span data-ttu-id="27d0c-107">Родительский элемент управления также может быть включен и отключен программным способом.</span><span class="sxs-lookup"><span data-stu-id="27d0c-107">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="27d0c-108">Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](../../design/disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="27d0c-108">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

## <a name="example"></a><span data-ttu-id="27d0c-109">Пример</span><span class="sxs-lookup"><span data-stu-id="27d0c-109">Example</span></span>

```xml
<Enabled>false</Enabled>
```
