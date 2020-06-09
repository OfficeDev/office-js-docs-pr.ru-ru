---
title: Элемент Hosts в файле манифеста
description: Указывает клиентское приложение Office, в котором будет активирована надстройка Office.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 037ac2b5fedbfb1b59b7523382574942fe59a00a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611808"
---
# <a name="hosts-element"></a><span data-ttu-id="a57e2-103">Элемент Hosts</span><span class="sxs-lookup"><span data-stu-id="a57e2-103">Hosts element</span></span>

<span data-ttu-id="a57e2-p101">Указывает клиентское приложение Office, в котором будет активирована надстройка Office. Содержит коллекцию элементов **Host** и их параметров.</span><span class="sxs-lookup"><span data-stu-id="a57e2-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="a57e2-106">При включении в узел [VersionOverrides](versionoverrides.md) этот элемент переопределяет элемент **Hosts** в родительской части манифеста.</span><span class="sxs-lookup"><span data-stu-id="a57e2-106">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="a57e2-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="a57e2-107">Child elements</span></span>

|  <span data-ttu-id="a57e2-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="a57e2-108">Element</span></span> |  <span data-ttu-id="a57e2-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="a57e2-109">Required</span></span>  |  <span data-ttu-id="a57e2-110">Описание</span><span class="sxs-lookup"><span data-stu-id="a57e2-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="a57e2-111">Host</span><span class="sxs-lookup"><span data-stu-id="a57e2-111">Host</span></span>](host.md)    |  <span data-ttu-id="a57e2-112">Да</span><span class="sxs-lookup"><span data-stu-id="a57e2-112">Yes</span></span>   |  <span data-ttu-id="a57e2-113">Описывает ведущее приложение и его параметры.</span><span class="sxs-lookup"><span data-stu-id="a57e2-113">Describes a host and its settings.</span></span> |
