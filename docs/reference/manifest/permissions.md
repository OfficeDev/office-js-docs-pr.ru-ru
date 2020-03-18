---
title: Элемент Permissions в файле манифеста
description: Элемент Permissions определяет уровень доступа к API для надстройки Office.
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 91e024a2f13ea7605941c8c17a642f325cbcd61d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718001"
---
# <a name="permissions-element"></a><span data-ttu-id="cd58e-103">Элемент Permissions</span><span class="sxs-lookup"><span data-stu-id="cd58e-103">Permissions element</span></span>

<span data-ttu-id="cd58e-104">Указывает уровень доступа к API для надстройки Office. Запрашивая разрешения, руководствуйтесь принципом минимальных привилегий.</span><span class="sxs-lookup"><span data-stu-id="cd58e-104">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="cd58e-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="cd58e-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="cd58e-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="cd58e-106">Syntax</span></span>

<span data-ttu-id="cd58e-107">Для надстроек области задач и контентных надстроек:</span><span class="sxs-lookup"><span data-stu-id="cd58e-107">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="cd58e-108">Для почтовых надстроек</span><span class="sxs-lookup"><span data-stu-id="cd58e-108">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="cd58e-109">Содержится в</span><span class="sxs-lookup"><span data-stu-id="cd58e-109">Contained in</span></span>

[<span data-ttu-id="cd58e-110">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="cd58e-110">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="cd58e-111">Замечания</span><span class="sxs-lookup"><span data-stu-id="cd58e-111">Remarks</span></span>

<span data-ttu-id="cd58e-112">Более подробную информацию можно узнать [в статье запрос разрешений для использования API в](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) надстройках и [Общие сведения о разрешениях для надстроек Outlook](../../outlook/understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="cd58e-112">For more detail, see [Requesting permissions for API use in add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).</span></span>
