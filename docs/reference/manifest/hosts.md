---
title: Элемент Hosts в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 59010c0f6c0d14d8721856f81def11540db28704
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433413"
---
# <a name="hosts-element"></a><span data-ttu-id="9ffa1-102">Элемент Hosts</span><span class="sxs-lookup"><span data-stu-id="9ffa1-102">Hosts element</span></span>

<span data-ttu-id="9ffa1-p101">Указывает клиентское приложение Office, в котором будет активирована надстройка Office. Содержит коллекцию элементов **Host** и их параметров.</span><span class="sxs-lookup"><span data-stu-id="9ffa1-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="9ffa1-105">При включении в узел [VersionOverrides](versionoverrides.md) этот элемент переопределяет элемент **Hosts** в родительской части манифеста.</span><span class="sxs-lookup"><span data-stu-id="9ffa1-105">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="9ffa1-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="9ffa1-106">Child elements</span></span>

|  <span data-ttu-id="9ffa1-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="9ffa1-107">Element</span></span> |  <span data-ttu-id="9ffa1-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="9ffa1-108">Required</span></span>  |  <span data-ttu-id="9ffa1-109">Описание</span><span class="sxs-lookup"><span data-stu-id="9ffa1-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="9ffa1-110">Host</span><span class="sxs-lookup"><span data-stu-id="9ffa1-110">Host</span></span>](host.md)    |  <span data-ttu-id="9ffa1-111">Да</span><span class="sxs-lookup"><span data-stu-id="9ffa1-111">Yes</span></span>   |  <span data-ttu-id="9ffa1-112">Описывает ведущее приложение и его параметры.</span><span class="sxs-lookup"><span data-stu-id="9ffa1-112">Describes a host and its settings.</span></span> |
