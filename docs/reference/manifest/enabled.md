---
title: Элемент Enabled в файле манифеста
description: Сведения о том, как указать, что команда надстройки отключена при запуске надстройки.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: a47ab97ff5a159c73bea52f130ce0c16efe2b6b6
ms.sourcegitcommit: 0e7ed44019d6564c79113639af831ea512fa0a13
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/09/2020
ms.locfileid: "42566215"
---
# <a name="enabled-element"></a><span data-ttu-id="ccfe6-103">Элемент Enabled</span><span class="sxs-lookup"><span data-stu-id="ccfe6-103">Enabled element</span></span>

<span data-ttu-id="ccfe6-104">Указывает, включен ли элемент управления ["Кнопка" или "](control.md#button-control) [меню](control.md#menu-dropdown-button-controls) " при запуске надстройки.</span><span class="sxs-lookup"><span data-stu-id="ccfe6-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="ccfe6-105">Элемент **Enabled** является дочерним элементом [элемента управления](control.md).</span><span class="sxs-lookup"><span data-stu-id="ccfe6-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="ccfe6-106">Если он не указан, используется `true`значение по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ccfe6-106">If it is omitted, the default is `true`.</span></span> 

<span data-ttu-id="ccfe6-107">Родительский элемент управления также может быть включен и отключен программным способом.</span><span class="sxs-lookup"><span data-stu-id="ccfe6-107">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="ccfe6-108">Дополнительные сведения можно найти [в статье Включение и отключение команд надстроек](/office/dev/add-ins/design/disable-add-in-commands).</span><span class="sxs-lookup"><span data-stu-id="ccfe6-108">For more information, see [Enable and Disable Add-in Commands](/office/dev/add-ins/design/disable-add-in-commands).</span></span>

## <a name="example"></a><span data-ttu-id="ccfe6-109">Пример</span><span class="sxs-lookup"><span data-stu-id="ccfe6-109">Example</span></span>

```xml
<Enabled>false</Enabled>
```

