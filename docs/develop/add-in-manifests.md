---
title: XML-манифест надстроек Office
description: ''
ms.date: 12/26/2018
ms.openlocfilehash: d4b83e0d450f98877169be16ea1cd391f2a6e9d6
ms.sourcegitcommit: 89e090a2a9ed85fe1b829820fb3b4dc6fb1fa6e9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/27/2018
ms.locfileid: "27447007"
---
# <a name="office-add-ins-xml-manifest"></a><span data-ttu-id="32ee5-102">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="32ee5-102">Office Add-ins XML manifest</span></span>

<span data-ttu-id="32ee5-103">XML-файл манифеста надстройки Office описывает способ ее активации, когда пользователь устанавливает и использует эту надстройку для работы с документами и приложениями Office.</span><span class="sxs-lookup"><span data-stu-id="32ee5-103">The XML manifest file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications.</span></span>

<span data-ttu-id="32ee5-104">С помощью такого XML-файла манифеста надстройка Office может выполнять следующие действия:</span><span class="sxs-lookup"><span data-stu-id="32ee5-104">An XML manifest file based on this schema enables an Office Add-in to do the following:</span></span>

* <span data-ttu-id="32ee5-105">предоставлять идентификатор, версию, описание, отображаемое имя и языковой стандарт по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="32ee5-105">Describe itself by providing an ID, version, description, display name, and default locale.</span></span>

* <span data-ttu-id="32ee5-106">указывать изображения, используемые для фирменного оформления надстройки, и значки, используемые для [команд надстройки][] в ленте Office;</span><span class="sxs-lookup"><span data-stu-id="32ee5-106">Specify the images used for branding the add-in and iconography used for [add-in commands][] in the Office Ribbon.</span></span>

* <span data-ttu-id="32ee5-107">указывать, как надстройка интегрируется с Office, включая создаваемые ею элементы пользовательского интерфейса, например кнопки на ленте;</span><span class="sxs-lookup"><span data-stu-id="32ee5-107">Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.</span></span>

* <span data-ttu-id="32ee5-108">определять запрошенные размеры по умолчанию для контентных надстроек, а также запрошенную высоту для надстроек Outlook;</span><span class="sxs-lookup"><span data-stu-id="32ee5-108">Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.</span></span>

* <span data-ttu-id="32ee5-109">объявлять разрешения, в которых нуждается надстройка Office, например чтение или запись документа;</span><span class="sxs-lookup"><span data-stu-id="32ee5-109">Declare permissions that the Office Add-in requires, such as reading or writing to the document.</span></span>

* <span data-ttu-id="32ee5-110">в случае надстроек Outlook необходимо определить одно или несколько правил, указывающих контекст, в котором эти надстройки будут активироваться и взаимодействовать с сообщением, сведениями о встрече или приглашением на собрание.</span><span class="sxs-lookup"><span data-stu-id="32ee5-110">For Outlook add-ins, define the rule or rules that specify the context in which they will be activated and interact with a message, appointment, or meeting request item.</span></span>

> [!NOTE]
> <span data-ttu-id="32ee5-p101">Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource и сделать ее доступной в интерфейсе Office, убедитесь, что она соответствует [политикам проверки AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и на [странице со сведениями о доступности и ведущих приложениях для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="32ee5-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="required-elements"></a><span data-ttu-id="32ee5-113">Обязательные элементы</span><span class="sxs-lookup"><span data-stu-id="32ee5-113">Required elements</span></span>

<span data-ttu-id="32ee5-114">В приведенной ниже таблице указаны обязательные элементы для трех типов надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="32ee5-114">The following table specifies the elements that are required for the three types of Office Add-ins.</span></span>

> [!NOTE]
> <span data-ttu-id="32ee5-115">Кроме того, есть обязательный порядок размещения элементов в родительском элементе.</span><span class="sxs-lookup"><span data-stu-id="32ee5-115">There is also a mandatory order in which elements must appear within their parent element.</span></span> <span data-ttu-id="32ee5-116">Дополнительные сведения см. в статье [Как определить правильный порядок элементов манифеста](manifest-element-ordering.md).</span><span class="sxs-lookup"><span data-stu-id="32ee5-116">For more information see [How to find the proper order of manifest elements](manifest-element-ordering.md).</span></span>


### <a name="required-elements-by-office-add-in-type"></a><span data-ttu-id="32ee5-117">Обязательные элементы по типам надстроек Office</span><span class="sxs-lookup"><span data-stu-id="32ee5-117">Required elements by Office Add-in type</span></span>

| <span data-ttu-id="32ee5-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="32ee5-118">Element</span></span>                                                                                      | <span data-ttu-id="32ee5-119">Контентная</span><span class="sxs-lookup"><span data-stu-id="32ee5-119">Content</span></span> | <span data-ttu-id="32ee5-120">Для области задач</span><span class="sxs-lookup"><span data-stu-id="32ee5-120">Task pane</span></span> | <span data-ttu-id="32ee5-121">Outlook</span><span class="sxs-lookup"><span data-stu-id="32ee5-121">Outlook</span></span> |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| <span data-ttu-id="32ee5-122">[OfficeApp][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-122">[OfficeApp][]</span></span>                                                                                |    <span data-ttu-id="32ee5-123">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-123">X</span></span>    |     <span data-ttu-id="32ee5-124">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-124">X</span></span>     |    <span data-ttu-id="32ee5-125">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-125">X</span></span>    |
| <span data-ttu-id="32ee5-126">
  [Id][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-126">[Id][]</span></span>                                                                                       |    <span data-ttu-id="32ee5-127">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-127">X</span></span>    |     <span data-ttu-id="32ee5-128">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-128">X</span></span>     |    <span data-ttu-id="32ee5-129">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-129">X</span></span>    |
| <span data-ttu-id="32ee5-130">
  [Version][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-130">[Version][]</span></span>                                                                                  |    <span data-ttu-id="32ee5-131">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-131">X</span></span>    |     <span data-ttu-id="32ee5-132">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-132">X</span></span>     |    <span data-ttu-id="32ee5-133">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-133">X</span></span>    |
| <span data-ttu-id="32ee5-134">[ProviderName][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-134">[ProviderName][]</span></span>                                                                             |    <span data-ttu-id="32ee5-135">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-135">X</span></span>    |     <span data-ttu-id="32ee5-136">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-136">X</span></span>     |    <span data-ttu-id="32ee5-137">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-137">X</span></span>    |
| <span data-ttu-id="32ee5-138">[DefaultLocale][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-138">[DefaultLocale][]</span></span>                                                                            |    <span data-ttu-id="32ee5-139">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-139">X</span></span>    |     <span data-ttu-id="32ee5-140">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-140">X</span></span>     |    <span data-ttu-id="32ee5-141">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-141">X</span></span>    |
| <span data-ttu-id="32ee5-142">[DisplayName][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-142">[DisplayName][]</span></span>                                                                              |    <span data-ttu-id="32ee5-143">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-143">X</span></span>    |     <span data-ttu-id="32ee5-144">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-144">X</span></span>     |    <span data-ttu-id="32ee5-145">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-145">X</span></span>    |
| <span data-ttu-id="32ee5-146">
  [Description][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-146">[Description][]</span></span>                                                                              |    <span data-ttu-id="32ee5-147">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-147">X</span></span>    |     <span data-ttu-id="32ee5-148">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-148">X</span></span>     |    <span data-ttu-id="32ee5-149">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-149">X</span></span>    |
| <span data-ttu-id="32ee5-150">[IconUrl][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-150">[IconUrl][]</span></span>                                                                                  |    <span data-ttu-id="32ee5-151">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-151">X</span></span>    |     <span data-ttu-id="32ee5-152">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-152">X</span></span>     |    <span data-ttu-id="32ee5-153">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-153">X</span></span>    |
| <span data-ttu-id="32ee5-154">[DefaultSettings (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-154">[DefaultSettings (ContentApp)][]</span></span><br/><span data-ttu-id="32ee5-155">[DefaultSettings (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-155">[DefaultSettings (TaskPaneApp)][]</span></span>                       |    <span data-ttu-id="32ee5-156">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-156">X</span></span>    |     <span data-ttu-id="32ee5-157">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-157">X</span></span>     |         |
| <span data-ttu-id="32ee5-158">[SourceLocation (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-158">[SourceLocation (ContentApp)][]</span></span><br/><span data-ttu-id="32ee5-159">[SourceLocation (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-159">[SourceLocation (TaskPaneApp)][]</span></span>                         |    <span data-ttu-id="32ee5-160">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-160">X</span></span>    |     <span data-ttu-id="32ee5-161">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-161">X</span></span>     |         |
| <span data-ttu-id="32ee5-162">[DesktopSettings][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-162">[DesktopSettings][]</span></span>                                                                          |         |           |    <span data-ttu-id="32ee5-163">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-163">X</span></span>    |
| <span data-ttu-id="32ee5-164">[SourceLocation (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-164">[SourceLocation (MailApp)][]</span></span>                                                                 |         |           |    <span data-ttu-id="32ee5-165">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-165">X</span></span>    |
| <span data-ttu-id="32ee5-166">
  [Permissions (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-166">[Permissions (ContentApp)][]</span></span><br/><span data-ttu-id="32ee5-167">
  [Permissions (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-167">[Permissions (TaskPaneApp)][]</span></span><br/><span data-ttu-id="32ee5-168">
  [Permissions (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-168">[Permissions (MailApp)][]</span></span> |    <span data-ttu-id="32ee5-169">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-169">X</span></span>    |     <span data-ttu-id="32ee5-170">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-170">X</span></span>     |    <span data-ttu-id="32ee5-171">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-171">X</span></span>    |
| <span data-ttu-id="32ee5-172">
  [Rule (RuleCollection)][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-172">[Rule (RuleCollection)][]</span></span><br/><span data-ttu-id="32ee5-173">
  [Rule (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-173">[Rule (MailApp)][]</span></span>                                             |         |           |    <span data-ttu-id="32ee5-174">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-174">X</span></span>    |
| <span data-ttu-id="32ee5-175">[Requirements (MailApp)\*][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-175">[Requirements (MailApp)\*][]</span></span>                                                                  |         |           |    <span data-ttu-id="32ee5-176">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-176">X</span></span>    |
| <span data-ttu-id="32ee5-177">[Set\*][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-177">[Set\*][]</span></span><br/><span data-ttu-id="32ee5-178">[Sets (MailAppRequirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-178">[Sets (MailAppRequirements)\*][]</span></span>                                                 |         |           |    <span data-ttu-id="32ee5-179">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-179">X</span></span>    |
| <span data-ttu-id="32ee5-180">[Form\*][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-180">[Form\*][]</span></span><br/><span data-ttu-id="32ee5-181">[FormSettings\*][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-181">[FormSettings\*][]</span></span>                                                              |         |           |    <span data-ttu-id="32ee5-182">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-182">X</span></span>    |
| <span data-ttu-id="32ee5-183">[Sets (Requirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-183">[Sets (Requirements)\*][]</span></span>                                                                     |    <span data-ttu-id="32ee5-184">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-184">X</span></span>    |     <span data-ttu-id="32ee5-185">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-185">X</span></span>     |         |
| <span data-ttu-id="32ee5-186">[Hosts\*][]</span><span class="sxs-lookup"><span data-stu-id="32ee5-186">[Hosts\*][]</span></span>                                                                                   |    <span data-ttu-id="32ee5-187">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-187">X</span></span>    |     <span data-ttu-id="32ee5-188">X</span><span class="sxs-lookup"><span data-stu-id="32ee5-188">X</span></span>     |         |

<span data-ttu-id="32ee5-189">_\*Элемент добавлен в схеме манифеста для надстроек Office версии 1.1._</span><span class="sxs-lookup"><span data-stu-id="32ee5-189">_\*Added in the Office Add-in Manifest Schema version 1.1._</span></span>

<!-- Links for above table -->

[officeapp]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/officeapp?view=office-js
[id]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/id
[version]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/version
[providername]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/providername
[defaultlocale]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/defaultlocale
[displayname]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/displayname
[description]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/description
[iconurl]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/iconurl
[defaultsettings (contentapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/defaultsettings
[defaultsettings (taskpaneapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/defaultsettings
[sourcelocation (contentapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation
[sourcelocation (taskpaneapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation
[desktopsettings]: https://msdn.microsoft.com/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c(Office.15).aspx
[sourcelocation (mailapp)]: https://msdn.microsoft.com/library/3792d389-bebd-d19a-9d90-35b7a0bfc623%28Office.15%29.aspx
[permissions (contentapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/permissions
[permissions (taskpaneapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/permissions
[permissions (mailapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/permissions
[rule (rulecollection)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/rule
[rule (mailapp)]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/rule
[requirements (mailapp)*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/requirements
[set*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/set
[sets (mailapprequirements)*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sets
[form*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/form
[formsettings*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/formsettings
[sets (requirements)*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sets
[hosts*]: https://docs.microsoft.com/office/dev/add-ins/reference/manifest/hosts

## <a name="hosting-requirements"></a><span data-ttu-id="32ee5-216">Требования к размещению</span><span class="sxs-lookup"><span data-stu-id="32ee5-216">Hosting requirements</span></span>

<span data-ttu-id="32ee5-217">Все URI изображений, в частности используемые для [команд надстройки][], должны поддерживать кэширование.</span><span class="sxs-lookup"><span data-stu-id="32ee5-217">All image URIs, such as those used for [add-in commands][], must support caching.</span></span> <span data-ttu-id="32ee5-218">Сервер с изображением не должен возвращать заголовок `Cache-Control`, содержащий `no-cache`, `no-store` или подобные параметры в ответе HTTP.</span><span class="sxs-lookup"><span data-stu-id="32ee5-218">The server hosting the image should not return a `Cache-Control` header specifying `no-cache`, `no-store`, or similar options in the HTTP response.</span></span>

<span data-ttu-id="32ee5-219">Все URL-адреса, например адреса исходных файлов, указанные в элементе [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation), должны быть **защищены с помощью SSL (HTTPS)**.</span><span class="sxs-lookup"><span data-stu-id="32ee5-219">All URLs, such as the source file locations specified in the [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) element, should be **SSL-secured (HTTPS)**.</span></span> [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a><span data-ttu-id="32ee5-220">Рекомендации по отправке решений в AppSource</span><span class="sxs-lookup"><span data-stu-id="32ee5-220">Best practices for submitting to AppSource</span></span>

<span data-ttu-id="32ee5-p104">Убедитесь, что идентификатор надстройки представляет собой допустимый и уникальный GUID. В Интернете доступно множество генераторов, с помощью которых можно создать уникальный GUID.</span><span class="sxs-lookup"><span data-stu-id="32ee5-p104">Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.</span></span>

<span data-ttu-id="32ee5-223">Надстройки, отправляемые в AppSource, также должны включать элемент [SupportUrl](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/supporturl).</span><span class="sxs-lookup"><span data-stu-id="32ee5-223">Add-ins submitted to AppSource must also include the [SupportUrl](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/supporturl) element.</span></span> <span data-ttu-id="32ee5-224">Дополнительные сведения см. в статье [Политики проверки для приложений и надстроек, отправляемых в AppSource](https://docs.microsoft.com/office/dev/store/validation-policies).</span><span class="sxs-lookup"><span data-stu-id="32ee5-224">For more information, see [Validation policies for apps and add-ins submitted to AppSource](https://docs.microsoft.com/office/dev/store/validation-policies).</span></span>

<span data-ttu-id="32ee5-225">Чтобы указать домены, отличные от указанного в элементе [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) для сценариев проверки подлинности, используйте только элемент [AppDomains](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomains).</span><span class="sxs-lookup"><span data-stu-id="32ee5-225">Only use the [AppDomains](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomains) element to specify domains other than the one specified in the [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) element for authentication scenarios.</span></span>

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a><span data-ttu-id="32ee5-226">Укажите домены, которые необходимо открыть в окне надстройки</span><span class="sxs-lookup"><span data-stu-id="32ee5-226">Specify domains you want to open in the add-in window</span></span>

<span data-ttu-id="32ee5-227">В Office Online область задач может открывать любой URL-адрес.</span><span class="sxs-lookup"><span data-stu-id="32ee5-227">When running in Office Online, your task pane can be navigated to any URL.</span></span> <span data-ttu-id="32ee5-228">Однако, на платформах для настольных компьютеров в области надстроек ведущего приложения Office открываются только URL-адреса в домене, где размещена начальная страница (указанная в элементе [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) файла манифеста).</span><span class="sxs-lookup"><span data-stu-id="32ee5-228">However, in desktop platforms, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/sourcelocation) element of the manifest file), that URL opens in a new browser window outside the add-in pane of the Office host application.</span></span>

<span data-ttu-id="32ee5-229">Чтобы переопределить это поведение, укажите все домены, которые должны открываться в окне надстройки, в списке доменов в элементе [AppDomains](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomains) файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="32ee5-229">To override this (desktop Office) behavior, specify each domain you want to open in the add-in window in the list of domains specified in the [AppDomains](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomains) element of the manifest file.</span></span> <span data-ttu-id="32ee5-230">URL-адреса в доменах из списка будут открываться в области задач и в классическом Office, и в Office Online.</span><span class="sxs-lookup"><span data-stu-id="32ee5-230">If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both desktop Office and Office Online.</span></span> <span data-ttu-id="32ee5-231">URL-адреса в доменах не из списка будут открываться в новом окне браузера (не в области надстроек) в классическом Office.</span><span class="sxs-lookup"><span data-stu-id="32ee5-231">If it tries to go to a URL that isn't in the list, then, in desktop Office, that URL opens in a new browser window (outside the add-in pane).</span></span>

> [!NOTE]
> <span data-ttu-id="32ee5-232">Это относится только к корневой области надстройки.</span><span class="sxs-lookup"><span data-stu-id="32ee5-232">This behavior applies only to the root pane of the add-in.</span></span> <span data-ttu-id="32ee5-233">Если в страницу надстройки внедрен iframe, его можно перенаправить на любой URL-адрес, независимо от того, указан ли он в элементе **AppDomains**, даже в классической версии Office.</span><span class="sxs-lookup"><span data-stu-id="32ee5-233">If there is an iframe embedded in the add-in page, the iframe can be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>

<span data-ttu-id="32ee5-234">В приведенном ниже примере XML-манифеста главная страница надстройки размещена в домене `https://www.contoso.com`, указанном в элементе **SourceLocation**.</span><span class="sxs-lookup"><span data-stu-id="32ee5-234">The following XML manifest example hosts its main add-in page in the `https://www.contoso.com` domain as specified in the **SourceLocation** element.</span></span> <span data-ttu-id="32ee5-235">В нем также указан домен `https://www.northwindtraders.com` с помощью элемента [AppDomain](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomain) из списка **AppDomains**.</span><span class="sxs-lookup"><span data-stu-id="32ee5-235">It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](https://docs.microsoft.com/office/dev/add-ins/reference/manifest/appdomain) element within the **AppDomains** element list.</span></span> <span data-ttu-id="32ee5-236">Страница в домене www.northwindtraders.com будет открываться в области надстроек даже в классической версии Office.</span><span class="sxs-lookup"><span data-stu-id="32ee5-236">If the add-in goes to a page in the www.northwindtraders.com domain, that page opens in the add-in pane, even in Office desktop.</span></span>

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="manifest-v11-xml-file-examples-and-schemas"></a><span data-ttu-id="32ee5-237">XML-файлы манифеста версии 1.1: примеры и схемы</span><span class="sxs-lookup"><span data-stu-id="32ee5-237">Manifest v1.1 XML file examples and schemas</span></span>

<span data-ttu-id="32ee5-238">Ниже показаны примеры XML-файлов манифеста версии 1.1 для надстроек области задач, контентных надстроек и надстроек Outlook.</span><span class="sxs-lookup"><span data-stu-id="32ee5-238">The following sections show examples of manifest v1.1 XML files for content, task pane, and Outlook add-ins.</span></span>

# <a name="task-panetabtabid-1"></a>[<span data-ttu-id="32ee5-239">Области задач</span><span class="sxs-lookup"><span data-stu-id="32ee5-239">Task pane</span></span>](#tab/tabid-1)

[<span data-ttu-id="32ee5-240">Схема манифеста приложения области задач</span><span class="sxs-lookup"><span data-stu-id="32ee5-240">Task pane app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/taskpane)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

  <!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

  <!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

  <!--IMPORTANT! Id must be unique for your add-in. If you clone this manifest ensure that you change this id to your own GUID -->
  <Id>e504fb41-a92a-4526-b101-542f357b7acb</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <!-- The display name of your add-in. Used on the store and various placed of the Office UI such as the add-ins dialog -->
  <DisplayName DefaultValue="Add-in Commands Sample" />
  <Description DefaultValue="Sample that illustrates add-in commands basic control types and actions" />
  <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />

  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />

  <!--BeginTaskpaneMode integration. Office 2013 and any client that doesn't understand commands will use this section.
    This section will also be used if there are no VersionOverrides -->
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
  </DefaultSettings>
  <!--EndTaskpaneMode integration -->

  <Permissions>ReadWriteDocument</Permissions>

  <!--BeginAddinCommandsMode integration-->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <!--Each host can have a different set of commands. Cool huh!? -->
      <!-- Workbook=Excel Document=Word Presentation=PowerPoint -->
      <!-- Make sure the hosts you override match the hosts declared in the top section of the manifest -->
      <Host xsi:type="Document">
        <!-- Form factor. Currently only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
        <DesktopFormFactor>
          <!--Function file is an html page that includes the javascript where functions for ExecuteAction will be called.
            Think of the FunctionFile as the "code behind" ExecuteFunction-->
          <FunctionFile resid="Contoso.FunctionFile.Url" />

          <!--PrimaryCommandSurface==Main Office Ribbon-->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!--Use OfficeTab to extend an existing Tab. Use CustomTab to create a new tab -->
            <!-- Documentation includes all the IDs currently tested to work -->
            <CustomTab id="Contoso.Tab1">
              <!--Group ID-->
              <Group id="Contoso.Tab1.Group1">
                <!--Label for your group. resid must point to a ShortString resource -->
                <Label resid="Contoso.Tab1.GroupLabel" />
                <Icon>
                  <!-- Sample Todo: Each size needs its own icon resource or it will look distorted when resized -->
                  <!--Icons. Required sizes 16,31,80, optional 20, 24, 40, 48, 64. Strongly recommended to provide all sizes for great UX -->
                  <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                </Icon>

                <!--Control. It can be of type "Button" or "Menu" -->
                <Control xsi:type="Button" id="Contoso.FunctionButton">
                  <!--Label for your button. resid must point to a ShortString resource -->
                  <Label resid="Contoso.FunctionButton.Label" />
                  <Supertip>
                    <!--ToolTip title. resid must point to a ShortString resource -->
                    <Title resid="Contoso.FunctionButton.Label" />
                    <!--ToolTip description. resid must point to a LongString resource -->
                    <Description resid="Contoso.FunctionButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon" />
                  </Icon>
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFunction or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  <Action xsi:type="ExecuteFunction">
                    <!--Name of the function to call. This function needs to exist in the global DOM namespace of the function file-->
                    <FunctionName>writeText</FunctionName>
                  </Action>
                </Control>

                <Control xsi:type="Button" id="Contoso.TaskpaneButton">
                  <Label resid="Contoso.TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="Contoso.TaskpaneButton.Label" />
                    <Description resid="Contoso.TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Button2Id1</TaskpaneId>
                    <!--Provide a url resource id for the location that will be displayed on the task pane -->
                    <SourceLocation resid="Contoso.Taskpane1.Url" />
                  </Action>
                </Control>
                <!-- Menu example -->
                <Control xsi:type="Menu" id="Contoso.Menu">
                  <Label resid="Contoso.Dropdown.Label" />
                  <Supertip>
                    <Title resid="Contoso.Dropdown.Label" />
                    <Description resid="Contoso.Dropdown.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                  </Icon>
                  <Items>
                    <Item id="Contoso.Menu.Item1">
                      <Label resid="Contoso.Item1.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item1.Label" />
                        <Description resid="Contoso.Item1.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID1</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane1.Url" />
                      </Action>
                    </Item>

                    <Item id="Contoso.Menu.Item2">
                      <Label resid="Contoso.Item2.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item2.Label" />
                        <Description resid="Contoso.Item2.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon" />
                      </Icon>
                      <Action xsi:type="ShowTaskpane">
                        <TaskpaneId>MyTaskPaneID2</TaskpaneId>
                        <SourceLocation resid="Contoso.Taskpane2.Url" />
                      </Action>
                    </Item>

                  </Items>
                </Control>

              </Group>

              <!-- Label of your tab -->
              <!-- If validating with XSD it needs to be at the end, we might change this before release -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon" DefaultValue="https://i.imgur.com/FkSShX9.png" />
        <bt:Image id="Contoso.FunctionButton.Icon" DefaultValue="https://i.imgur.com/qDujiX0.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Contoso.FunctionFile.Url" DefaultValue="https://commandsimple.azurewebsites.net/FunctionFile.html" />
        <bt:Url id="Contoso.Taskpane1.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane.html" />
        <bt:Url id="Contoso.Taskpane2.Url" DefaultValue="https://commandsimple.azurewebsites.net/Taskpane2.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Contoso.FunctionButton.Label" DefaultValue="Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Label" DefaultValue="Show Taskpane" />
        <bt:String id="Contoso.Dropdown.Label" DefaultValue="Dropdown" />
        <bt:String id="Contoso.Item1.Label" DefaultValue="Show Taskpane 1" />
        <bt:String id="Contoso.Item2.Label" DefaultValue="Show Taskpane 2" />
        <bt:String id="Contoso.Tab1.GroupLabel" DefaultValue="Test Group" />
         <bt:String id="Contoso.Tab1.TabLabel" DefaultValue="Test Tab" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Contoso.FunctionButton.Tooltip" DefaultValue="Click to Execute Function" />
        <bt:String id="Contoso.TaskpaneButton.Tooltip" DefaultValue="Click to Show a Taskpane" />
        <bt:String id="Contoso.Dropdown.Tooltip" DefaultValue="Click to Show Options on this Menu" />
        <bt:String id="Contoso.Item1.Tooltip" DefaultValue="Click to Show Taskpane1" />
        <bt:String id="Contoso.Item2.Tooltip" DefaultValue="Click to Show Taskpane2" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

# <a name="contenttabtabid-2"></a>[<span data-ttu-id="32ee5-241">Контентная</span><span class="sxs-lookup"><span data-stu-id="32ee5-241">Content</span></span>](#tab/tabid-2)

[<span data-ttu-id="32ee5-242">Схема манифеста контентного приложения</span><span class="sxs-lookup"><span data-stu-id="32ee5-242">Content app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/content)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <Id>01eac144-e55a-45a7-b6e3-f1cc60ab0126</Id>
  <AlternateId>en-US\WA123456789</AlternateId>
  <Version>1.0.0.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Sample content add-in" />
  <Description DefaultValue="Describe the features of this app." />
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />

  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />

  <Hosts>
    <Host Name="Workbook" />
    <Host Name="Database" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="TableBindings" />
    </Sets>
  </Requirements>  
  <DefaultSettings>
    <SourceLocation DefaultValue="https://contoso.com/apps/content.html" />
    <RequestedWidth>400</RequestedWidth>
    <RequestedHeight>400</RequestedHeight>
  </DefaultSettings>
  <Permissions>Restricted</Permissions>
  <AllowSnapshot>true</AllowSnapshot>
</OfficeApp>
```

# <a name="mailtabtabid-3"></a>[<span data-ttu-id="32ee5-243">Почтовая</span><span class="sxs-lookup"><span data-stu-id="32ee5-243">Mail</span></span>](#tab/tabid-3)

[<span data-ttu-id="32ee5-244">Схема манифеста почтового приложения</span><span class="sxs-lookup"><span data-stu-id="32ee5-244">Mail app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/mail)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">

  <Id>971E76EF-D73E-567F-ADAE-5A76B39052CF</Id>
  <Version>1.0</Version>
  <ProviderName>Microsoft</ProviderName>
  <DefaultLocale>en-us</DefaultLocale>
  <DisplayName DefaultValue="YouTube"/>
  <Description DefaultValue=
    "Watch YouTube videos referenced in the e-mails you  
    receive without leaving your email client.">
    <Override Locale="fr-fr" Value="Visualisez les vidéos
      YouTube références dans vos courriers électronique
      directement depuis Outlook et Outlook Web App."/>
  </Description>
  <!-- Change the following lines to specify    -->
  <!-- the web server that hosts the icon files. -->
  <IconUrl DefaultValue="https://contoso.com/assets/icon-64.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png" />

  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />

  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>

  <FormSettings>
    <Form xsi:type="ItemRead">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_desktop.htm" />
        <RequestedHeight>216</RequestedHeight>
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_read_tablet.htm" />
        <RequestedHeight>216</RequestedHeight>
      </TabletSettings>
    </Form>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_desktop.htm" />
      </DesktopSettings>
      <TabletSettings>
        <!-- Change the following line to specify     -->
        <!-- the web server that hosts the HTML file. -->
        <SourceLocation DefaultValue=
          "https://webserver/YouTube/YouTube_compose_tablet.htm" />
      </TabletSettings>
    </Form>
  </FormSettings>

  <Permissions>ReadWriteItem</Permissions>
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="RuleCollection" Mode="And">
      <Rule xsi:type="RuleCollection" Mode="Or">
        <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
        <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
      </Rule>
      <Rule xsi:type="ItemHasRegularExpressionMatch"
        PropertyName="BodyAsPlaintext" RegExName="VideoURL"
        RegExValue=
        "http://(((www\.)?youtube\.com/watch\?v=)|
        (youtu\.be/))[a-zA-Z0-9_-]{11}" />
    </Rule>
    <Rule xsi:type="RuleCollection" Mode="Or">
      <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit" />
      <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
    </Rule>
  </Rule>
</OfficeApp>
```

---

## <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="32ee5-245">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="32ee5-245">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="32ee5-p110">Сведения об устранении проблем, связанных с манифестом надстройки, см. в статье [Проверка манифеста и устранение связанных с ним неполадок](../testing/troubleshoot-manifest.md). Там указано, как проверить манифест согласно [XSD](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), а также как отладить манифест с помощью ведения журнала в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="32ee5-p110">For troubleshooting issues with your manifest, see [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md). There, you will find information on how to validate the manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), and also how to use runtime logging to debug the manifest.</span></span>

## <a name="see-also"></a><span data-ttu-id="32ee5-248">См. также</span><span class="sxs-lookup"><span data-stu-id="32ee5-248">See also</span></span>

* <span data-ttu-id="32ee5-249">[Создание команд надстройки в манифесте][команды надстройки]</span><span class="sxs-lookup"><span data-stu-id="32ee5-249">[Create add-in commands in your manifest][add-in commands]</span></span>
* [<span data-ttu-id="32ee5-250">Указание ведущих приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="32ee5-250">Specify Office hosts and API requirements</span></span>](specify-office-hosts-and-api-requirements.md)
* [<span data-ttu-id="32ee5-251">Локализация надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="32ee5-251">Localization for Office Add-ins</span></span>](localization.md)
* [<span data-ttu-id="32ee5-252">Справочная схема по манифестам надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="32ee5-252">Schema reference for Office Add-ins manifests</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)
* [<span data-ttu-id="32ee5-253">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="32ee5-253">Validate and troubleshoot issues with your manifest</span></span>](../testing/troubleshoot-manifest.md)

[команды надстройки]: create-addin-commands.md
[add-in commands]: create-addin-commands.md