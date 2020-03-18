---
title: Элемент Hosts в файле манифеста
description: Указывает клиентское приложение Office, в котором будет активирована надстройка Office.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cd4e0eecce610b10fdc9dafcde7b807fde425b14
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718106"
---
# <a name="hosts-element"></a><span data-ttu-id="ab46f-103">Элемент Hosts</span><span class="sxs-lookup"><span data-stu-id="ab46f-103">Hosts element</span></span>

<span data-ttu-id="ab46f-p101">Указывает клиентское приложение Office, в котором будет активирована надстройка Office. Содержит коллекцию элементов **Host** и их параметров.</span><span class="sxs-lookup"><span data-stu-id="ab46f-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="ab46f-106">При включении в узел [VersionOverrides](versionoverrides.md) этот элемент переопределяет элемент **Hosts** в родительской части манифеста.</span><span class="sxs-lookup"><span data-stu-id="ab46f-106">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="ab46f-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="ab46f-107">Child elements</span></span>

|  <span data-ttu-id="ab46f-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="ab46f-108">Element</span></span> |  <span data-ttu-id="ab46f-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="ab46f-109">Required</span></span>  |  <span data-ttu-id="ab46f-110">Описание</span><span class="sxs-lookup"><span data-stu-id="ab46f-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ab46f-111">Host</span><span class="sxs-lookup"><span data-stu-id="ab46f-111">Host</span></span>](host.md)    |  <span data-ttu-id="ab46f-112">Да</span><span class="sxs-lookup"><span data-stu-id="ab46f-112">Yes</span></span>   |  <span data-ttu-id="ab46f-113">Описывает ведущее приложение и его параметры.</span><span class="sxs-lookup"><span data-stu-id="ab46f-113">Describes a host and its settings.</span></span> |
