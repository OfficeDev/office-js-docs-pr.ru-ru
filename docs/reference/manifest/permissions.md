---
title: Элемент Permissions в файле манифеста
description: Элемент Permissions определяет уровень доступа к API для надстройки Office.
ms.date: 06/26/2020
localization_priority: Normal
ms.openlocfilehash: bc4cc2713d5a781c3407385470acd762910d17fd
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006460"
---
# <a name="permissions-element"></a><span data-ttu-id="4c439-103">Элемент Permissions</span><span class="sxs-lookup"><span data-stu-id="4c439-103">Permissions element</span></span>

<span data-ttu-id="4c439-104">Указывает уровень доступа к API для надстройки Office. Запрашивая разрешения, руководствуйтесь принципом минимальных привилегий.</span><span class="sxs-lookup"><span data-stu-id="4c439-104">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="4c439-105">**Тип надстройки:** контентные и почтовые надстройки, надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="4c439-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4c439-106">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="4c439-106">Syntax</span></span>

<span data-ttu-id="4c439-107">Для надстроек области задач и контентных надстроек:</span><span class="sxs-lookup"><span data-stu-id="4c439-107">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="4c439-108">Для почтовых надстроек:</span><span class="sxs-lookup"><span data-stu-id="4c439-108">For mail add-ins:</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="4c439-109">Содержится в</span><span class="sxs-lookup"><span data-stu-id="4c439-109">Contained in</span></span>

[<span data-ttu-id="4c439-110">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="4c439-110">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="4c439-111">Замечания</span><span class="sxs-lookup"><span data-stu-id="4c439-111">Remarks</span></span>

<span data-ttu-id="4c439-112">Дополнительные сведения см в статье [запрашивание разрешений для использования API в контентных](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) надстройках и надстройках области задач и [Общие сведения о разрешениях для надстроек Outlook](../../outlook/understanding-outlook-add-in-permissions.md).</span><span class="sxs-lookup"><span data-stu-id="4c439-112">For more details, see [Requesting permissions for API use in content and task pane add-ins](../../develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md) and [Understanding Outlook add-in permissions](../../outlook/understanding-outlook-add-in-permissions.md).</span></span>
