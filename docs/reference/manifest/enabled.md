---
title: Элемент Enabled в файле манифеста
description: Узнайте, как указать, что команда надстройки отключена при запуске надстройки.
ms.date: 01/04/2021
localization_priority: Normal
ms.openlocfilehash: be18767638af6f2be6352cea46739f6a01b7dd45
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771399"
---
# <a name="enabled-element"></a><span data-ttu-id="9df0b-103">Элемент Enabled</span><span class="sxs-lookup"><span data-stu-id="9df0b-103">Enabled element</span></span>

<span data-ttu-id="9df0b-104">Указывает, включен ли [](control.md#menu-dropdown-button-controls) [при](control.md#button-control) запуске надстройки пункт "Кнопка" или "Меню".</span><span class="sxs-lookup"><span data-stu-id="9df0b-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="9df0b-105">Элемент **Enabled** — это элемент [control.](control.md)</span><span class="sxs-lookup"><span data-stu-id="9df0b-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="9df0b-106">Если он опущен, значение по умолчанию `true` : .</span><span class="sxs-lookup"><span data-stu-id="9df0b-106">If it is omitted, the default is `true`.</span></span>

<span data-ttu-id="9df0b-107">Этот элемент действителен только в Excel; то есть, если `Name` атрибутом элемента [Host](host.md) является "Workbook".</span><span class="sxs-lookup"><span data-stu-id="9df0b-107">This element is only valid in Excel; that is, when the `Name` attribute of the [Host](host.md) element is "Workbook".</span></span>

<span data-ttu-id="9df0b-108">Родительский контроль также можно включить программным образом и отключить.</span><span class="sxs-lookup"><span data-stu-id="9df0b-108">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="9df0b-109">Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](../../design/disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="9df0b-109">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

## <a name="example"></a><span data-ttu-id="9df0b-110">Пример</span><span class="sxs-lookup"><span data-stu-id="9df0b-110">Example</span></span>

```xml
<Enabled>false</Enabled>
```
