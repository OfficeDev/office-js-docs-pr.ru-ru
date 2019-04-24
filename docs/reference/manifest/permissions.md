---
title: Элемент Permissions в файле манифеста
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3442a8e0caee442ce1b38c5ff39cfd1ef5088fb7
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450662"
---
# <a name="permissions-element"></a><span data-ttu-id="84eb0-102">Элемент Permissions</span><span class="sxs-lookup"><span data-stu-id="84eb0-102">Permissions element</span></span>

<span data-ttu-id="84eb0-103">Указывает уровень доступа к API для надстройки Office. Запрашивая разрешения, руководствуйтесь принципом минимальных привилегий.</span><span class="sxs-lookup"><span data-stu-id="84eb0-103">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="84eb0-104">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="84eb0-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="84eb0-105">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="84eb0-105">Syntax</span></span>

<span data-ttu-id="84eb0-106">Для надстроек области задач и контентных надстроек:</span><span class="sxs-lookup"><span data-stu-id="84eb0-106">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="84eb0-107">Для почтовых надстроек</span><span class="sxs-lookup"><span data-stu-id="84eb0-107">For mail add-ins</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="84eb0-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="84eb0-108">Contained in</span></span>

[<span data-ttu-id="84eb0-109">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="84eb0-109">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="84eb0-110">Замечания</span><span class="sxs-lookup"><span data-stu-id="84eb0-110">Remarks</span></span>

<span data-ttu-id="84eb0-111">Подробные сведения см. в статьях [Запрашивание разрешений на использование API в надстройках области задач и контентных надстройках](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) и [Общие сведения о разрешениях для надстроек Outlook](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="84eb0-111">For more detail, see [Requesting permissions for API use in content and task pane add-ins](/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
