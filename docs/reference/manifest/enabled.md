---
title: Элемент Enabled в файле манифеста
description: Сведения о том, как указать, что команда надстройки отключена при запуске надстройки.
ms.date: 01/10/2020
localization_priority: Normal
ms.openlocfilehash: 2849689fec99190c3a9b039c6c04069bc8194ee1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611570"
---
# <a name="enabled-element"></a><span data-ttu-id="fc147-103">Элемент Enabled</span><span class="sxs-lookup"><span data-stu-id="fc147-103">Enabled element</span></span>

<span data-ttu-id="fc147-104">Указывает, включен ли элемент управления ["Кнопка" или "](control.md#button-control) [меню](control.md#menu-dropdown-button-controls) " при запуске надстройки.</span><span class="sxs-lookup"><span data-stu-id="fc147-104">Specifies whether a [Button](control.md#button-control) or [Menu](control.md#menu-dropdown-button-controls) control is enabled when the add-in launches.</span></span> <span data-ttu-id="fc147-105">Элемент **Enabled** является дочерним элементом [элемента управления](control.md).</span><span class="sxs-lookup"><span data-stu-id="fc147-105">The **Enabled** element is a child element of [Control](control.md).</span></span> <span data-ttu-id="fc147-106">Если он не указан, используется значение по умолчанию `true` .</span><span class="sxs-lookup"><span data-stu-id="fc147-106">If it is omitted, the default is `true`.</span></span>

<span data-ttu-id="fc147-107">Родительский элемент управления также может быть включен и отключен программным способом.</span><span class="sxs-lookup"><span data-stu-id="fc147-107">The parent control can also be programmatically enabled and disabled.</span></span> <span data-ttu-id="fc147-108">Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](../../design/disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="fc147-108">For more information, see [Enable and Disable Add-in Commands](../../design/disable-add-in-commands.md).</span></span>

## <a name="example"></a><span data-ttu-id="fc147-109">Пример</span><span class="sxs-lookup"><span data-stu-id="fc147-109">Example</span></span>

```xml
<Enabled>false</Enabled>
```
