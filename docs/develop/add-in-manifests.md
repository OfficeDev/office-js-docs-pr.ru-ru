---
title: XML-манифест надстроек Office
description: Получите обзор манифеста надстройки Office и его использования.
ms.date: 03/18/2020
localization_priority: Priority
ms.openlocfilehash: 495638ee70630c5330e800419076463273bd2491
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293356"
---
# <a name="office-add-ins-xml-manifest"></a><span data-ttu-id="ba40f-103">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="ba40f-103">Office Add-ins XML manifest</span></span>

<span data-ttu-id="ba40f-104">XML-файл манифеста надстройки Office описывает способ ее активации, когда пользователь устанавливает и использует эту надстройку для работы с документами и приложениями Office.</span><span class="sxs-lookup"><span data-stu-id="ba40f-104">The XML manifest file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications.</span></span>

<span data-ttu-id="ba40f-105">С помощью такого XML-файла манифеста надстройка Office может выполнять следующие действия:</span><span class="sxs-lookup"><span data-stu-id="ba40f-105">An XML manifest file based on this schema enables an Office Add-in to do the following:</span></span>

* <span data-ttu-id="ba40f-106">предоставлять идентификатор, версию, описание, отображаемое имя и языковой стандарт по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ba40f-106">Describe itself by providing an ID, version, description, display name, and default locale.</span></span>

* <span data-ttu-id="ba40f-107">указывать изображения, используемые для фирменного оформления надстройки, и значки, используемые для [команд надстройки][] на ленте приложения Office;</span><span class="sxs-lookup"><span data-stu-id="ba40f-107">Specify the images used for branding the add-in and iconography used for [add-in commands][] in the Office app ribbon.</span></span>

* <span data-ttu-id="ba40f-108">указывать, как надстройка интегрируется с Office, включая создаваемые ею элементы пользовательского интерфейса, например кнопки на ленте;</span><span class="sxs-lookup"><span data-stu-id="ba40f-108">Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.</span></span>

* <span data-ttu-id="ba40f-109">определять запрошенные размеры по умолчанию для контентных надстроек, а также запрошенную высоту для надстроек Outlook;</span><span class="sxs-lookup"><span data-stu-id="ba40f-109">Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.</span></span>

* <span data-ttu-id="ba40f-110">объявлять разрешения, в которых нуждается надстройка Office, например чтение или запись документа;</span><span class="sxs-lookup"><span data-stu-id="ba40f-110">Declare permissions that the Office Add-in requires, such as reading or writing to the document.</span></span>

* <span data-ttu-id="ba40f-111">В случае надстроек Outlook необходимо определить одно или несколько правил, указывающих контекст, в котором эти надстройки будут активироваться и взаимодействовать с сообщением, сведениями о встрече или приглашением на собрание.</span><span class="sxs-lookup"><span data-stu-id="ba40f-111">For Outlook add-ins, define the rule or rules that specify the context in which they will be activated and interact with a message, appointment, or meeting request item.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="required-elements"></a><span data-ttu-id="ba40f-112">Обязательные элементы</span><span class="sxs-lookup"><span data-stu-id="ba40f-112">Required elements</span></span>

<span data-ttu-id="ba40f-113">В приведенной ниже таблице указаны обязательные элементы для трех типов надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="ba40f-113">The following table specifies the elements that are required for the three types of Office Add-ins.</span></span>

> [!NOTE]
> <span data-ttu-id="ba40f-114">Кроме того, есть обязательный порядок размещения элементов в родительском элементе.</span><span class="sxs-lookup"><span data-stu-id="ba40f-114">There is also a mandatory order in which elements must appear within their parent element.</span></span> <span data-ttu-id="ba40f-115">Дополнительные сведения см. в статье [Как определить правильный порядок элементов манифеста](manifest-element-ordering.md).</span><span class="sxs-lookup"><span data-stu-id="ba40f-115">For more information see [How to find the proper order of manifest elements](manifest-element-ordering.md).</span></span>


### <a name="required-elements-by-office-add-in-type"></a><span data-ttu-id="ba40f-116">Обязательные элементы по типам надстроек Office</span><span class="sxs-lookup"><span data-stu-id="ba40f-116">Required elements by Office Add-in type</span></span>

| <span data-ttu-id="ba40f-117">Элемент</span><span class="sxs-lookup"><span data-stu-id="ba40f-117">Element</span></span>                                                                                      | <span data-ttu-id="ba40f-118">Контентная</span><span class="sxs-lookup"><span data-stu-id="ba40f-118">Content</span></span> | <span data-ttu-id="ba40f-119">Для области задач</span><span class="sxs-lookup"><span data-stu-id="ba40f-119">Task pane</span></span> | <span data-ttu-id="ba40f-120">Outlook</span><span class="sxs-lookup"><span data-stu-id="ba40f-120">Outlook</span></span> |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| <span data-ttu-id="ba40f-121">[OfficeApp][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-121">[OfficeApp][]</span></span>                                                                                |    <span data-ttu-id="ba40f-122">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-122">X</span></span>    |     <span data-ttu-id="ba40f-123">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-123">X</span></span>     |    <span data-ttu-id="ba40f-124">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-124">X</span></span>    |
| <span data-ttu-id="ba40f-125">
  [Id][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-125">[Id][]</span></span>                                                                                       |    <span data-ttu-id="ba40f-126">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-126">X</span></span>    |     <span data-ttu-id="ba40f-127">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-127">X</span></span>     |    <span data-ttu-id="ba40f-128">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-128">X</span></span>    |
| <span data-ttu-id="ba40f-129">
  [Version][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-129">[Version][]</span></span>                                                                                  |    <span data-ttu-id="ba40f-130">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-130">X</span></span>    |     <span data-ttu-id="ba40f-131">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-131">X</span></span>     |    <span data-ttu-id="ba40f-132">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-132">X</span></span>    |
| <span data-ttu-id="ba40f-133">[ProviderName][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-133">[ProviderName][]</span></span>                                                                             |    <span data-ttu-id="ba40f-134">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-134">X</span></span>    |     <span data-ttu-id="ba40f-135">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-135">X</span></span>     |    <span data-ttu-id="ba40f-136">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-136">X</span></span>    |
| <span data-ttu-id="ba40f-137">[DefaultLocale][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-137">[DefaultLocale][]</span></span>                                                                            |    <span data-ttu-id="ba40f-138">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-138">X</span></span>    |     <span data-ttu-id="ba40f-139">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-139">X</span></span>     |    <span data-ttu-id="ba40f-140">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-140">X</span></span>    |
| <span data-ttu-id="ba40f-141">[DisplayName][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-141">[DisplayName][]</span></span>                                                                              |    <span data-ttu-id="ba40f-142">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-142">X</span></span>    |     <span data-ttu-id="ba40f-143">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-143">X</span></span>     |    <span data-ttu-id="ba40f-144">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-144">X</span></span>    |
| <span data-ttu-id="ba40f-145">
  [Description][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-145">[Description][]</span></span>                                                                              |    <span data-ttu-id="ba40f-146">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-146">X</span></span>    |     <span data-ttu-id="ba40f-147">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-147">X</span></span>     |    <span data-ttu-id="ba40f-148">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-148">X</span></span>    |
| <span data-ttu-id="ba40f-149">[IconUrl][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-149">[IconUrl][]</span></span>                                                                                  |    <span data-ttu-id="ba40f-150">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-150">X</span></span>    |     <span data-ttu-id="ba40f-151">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-151">X</span></span>     |    <span data-ttu-id="ba40f-152">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-152">X</span></span>    |
| <span data-ttu-id="ba40f-153">[SupportUrl][]\*\*</span><span class="sxs-lookup"><span data-stu-id="ba40f-153">[SupportUrl][]\*\*</span></span>                                                                           |    <span data-ttu-id="ba40f-154">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-154">X</span></span>    |     <span data-ttu-id="ba40f-155">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-155">X</span></span>     |    <span data-ttu-id="ba40f-156">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-156">X</span></span>    |
| <span data-ttu-id="ba40f-157">[DefaultSettings (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-157">[DefaultSettings (ContentApp)][]</span></span><br/><span data-ttu-id="ba40f-158">[DefaultSettings (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-158">[DefaultSettings (TaskPaneApp)][]</span></span>                       |    <span data-ttu-id="ba40f-159">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-159">X</span></span>    |     <span data-ttu-id="ba40f-160">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-160">X</span></span>     |         |
| <span data-ttu-id="ba40f-161">[SourceLocation (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-161">[SourceLocation (ContentApp)][]</span></span><br/><span data-ttu-id="ba40f-162">[SourceLocation (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-162">[SourceLocation (TaskPaneApp)][]</span></span>                         |    <span data-ttu-id="ba40f-163">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-163">X</span></span>    |     <span data-ttu-id="ba40f-164">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-164">X</span></span>     |         |
| <span data-ttu-id="ba40f-165">[DesktopSettings][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-165">[DesktopSettings][]</span></span>                                                                          |         |           |    <span data-ttu-id="ba40f-166">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-166">X</span></span>    |
| <span data-ttu-id="ba40f-167">[SourceLocation (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-167">[SourceLocation (MailApp)][]</span></span>                                                                 |         |           |    <span data-ttu-id="ba40f-168">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-168">X</span></span>    |
| <span data-ttu-id="ba40f-169">
  [Permissions (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-169">[Permissions (ContentApp)][]</span></span><br/><span data-ttu-id="ba40f-170">
  [Permissions (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-170">[Permissions (TaskPaneApp)][]</span></span><br/><span data-ttu-id="ba40f-171">
  [Permissions (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-171">[Permissions (MailApp)][]</span></span> |    <span data-ttu-id="ba40f-172">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-172">X</span></span>    |     <span data-ttu-id="ba40f-173">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-173">X</span></span>     |    <span data-ttu-id="ba40f-174">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-174">X</span></span>    |
| <span data-ttu-id="ba40f-175">
  [Rule (RuleCollection)][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-175">[Rule (RuleCollection)][]</span></span><br/><span data-ttu-id="ba40f-176">
  [Rule (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-176">[Rule (MailApp)][]</span></span>                                             |         |           |    <span data-ttu-id="ba40f-177">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-177">X</span></span>    |
| <span data-ttu-id="ba40f-178">[Requirements (MailApp)\*][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-178">[Requirements (MailApp)\*][]</span></span>                                                                  |         |           |    <span data-ttu-id="ba40f-179">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-179">X</span></span>    |
| <span data-ttu-id="ba40f-180">[Set\*][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-180">[Set\*][]</span></span><br/><span data-ttu-id="ba40f-181">[Sets (MailAppRequirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-181">[Sets (MailAppRequirements)\*][]</span></span>                                                 |         |           |    <span data-ttu-id="ba40f-182">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-182">X</span></span>    |
| <span data-ttu-id="ba40f-183">[Form\*][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-183">[Form\*][]</span></span><br/><span data-ttu-id="ba40f-184">[FormSettings\*][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-184">[FormSettings\*][]</span></span>                                                              |         |           |    <span data-ttu-id="ba40f-185">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-185">X</span></span>    |
| <span data-ttu-id="ba40f-186">[Sets (Requirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-186">[Sets (Requirements)\*][]</span></span>                                                                     |    <span data-ttu-id="ba40f-187">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-187">X</span></span>    |     <span data-ttu-id="ba40f-188">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-188">X</span></span>     |         |
| <span data-ttu-id="ba40f-189">[Hosts\*][]</span><span class="sxs-lookup"><span data-stu-id="ba40f-189">[Hosts\*][]</span></span>                                                                                   |    <span data-ttu-id="ba40f-190">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-190">X</span></span>    |     <span data-ttu-id="ba40f-191">X</span><span class="sxs-lookup"><span data-stu-id="ba40f-191">X</span></span>     |         |

<span data-ttu-id="ba40f-192">_\*Элемент добавлен в схеме манифеста для надстроек Office версии 1.1._</span><span class="sxs-lookup"><span data-stu-id="ba40f-192">_\*Added in the Office Add-in Manifest Schema version 1.1._</span></span>

<span data-ttu-id="ba40f-193">_\*\* SupportUrl требуется только для надстроек распространяемых с помощью AppSource._</span><span class="sxs-lookup"><span data-stu-id="ba40f-193">_\*\* SupportUrl is only required for add-ins that are distributed through AppSource._</span></span>

<!-- Links for above table -->

[officeapp]: ../reference/manifest/officeapp.md
[id]: ../reference/manifest/id.md
[version]: ../reference/manifest/version.md
[providername]: ../reference/manifest/providername.md
[defaultlocale]: ../reference/manifest/defaultlocale.md
[displayname]: ../reference/manifest/displayname.md
[description]: ../reference/manifest/description.md
[iconurl]: ../reference/manifest/iconurl.md
[supporturl]: ../reference/manifest/supporturl.md
[defaultsettings (contentapp)]: ../reference/manifest/defaultsettings.md
[defaultsettings (taskpaneapp)]: ../reference/manifest/defaultsettings.md
[sourcelocation (contentapp)]: ../reference/manifest/sourcelocation.md
[sourcelocation (taskpaneapp)]: ../reference/manifest/sourcelocation.md
[desktopsettings]: /previous-versions/office/fp179684%28v=office.15%29
[sourcelocation (mailapp)]: /previous-versions/office/fp123668%28v=office.15%29
[permissions (contentapp)]: ../reference/manifest/permissions.md
[permissions (taskpaneapp)]: ../reference/manifest/permissions.md
[permissions (mailapp)]: ../reference/manifest/permissions.md
[rule (rulecollection)]: ../reference/manifest/rule.md
[rule (mailapp)]: ../reference/manifest/rule.md
[requirements (mailapp)*]: ../reference/manifest/requirements.md
[set*]: ../reference/manifest/set.md
[sets (mailapprequirements)*]: ../reference/manifest/sets.md
[form*]: ../reference/manifest/form.md
[formsettings*]: ../reference/manifest/formsettings.md
[sets (requirements)*]: ../reference/manifest/sets.md
[hosts*]: ../reference/manifest/hosts.md

## <a name="hosting-requirements"></a><span data-ttu-id="ba40f-221">Требования к размещению</span><span class="sxs-lookup"><span data-stu-id="ba40f-221">Hosting requirements</span></span>

<span data-ttu-id="ba40f-222">Все URI изображений, в частности используемые для [команд надстройки][], должны поддерживать кэширование.</span><span class="sxs-lookup"><span data-stu-id="ba40f-222">All image URIs, such as those used for [add-in commands][], must support caching.</span></span> <span data-ttu-id="ba40f-223">Сервер с изображением не должен возвращать заголовок `Cache-Control`, содержащий `no-cache`, `no-store` или подобные параметры в ответе HTTP.</span><span class="sxs-lookup"><span data-stu-id="ba40f-223">The server hosting the image should not return a `Cache-Control` header specifying `no-cache`, `no-store`, or similar options in the HTTP response.</span></span>

<span data-ttu-id="ba40f-224">Все URL-адреса, например адреса исходных файлов, указанные в элементе [SourceLocation](../reference/manifest/sourcelocation.md), должны быть **защищены с помощью SSL (HTTPS)**.</span><span class="sxs-lookup"><span data-stu-id="ba40f-224">All URLs, such as the source file locations specified in the [SourceLocation](../reference/manifest/sourcelocation.md) element, should be **SSL-secured (HTTPS)**.</span></span> [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a><span data-ttu-id="ba40f-225">Рекомендации по отправке решений в AppSource</span><span class="sxs-lookup"><span data-stu-id="ba40f-225">Best practices for submitting to AppSource</span></span>

<span data-ttu-id="ba40f-p103">Убедитесь, что идентификатор надстройки представляет собой допустимый и уникальный GUID. В Интернете доступно множество генераторов, с помощью которых можно создать уникальный GUID.</span><span class="sxs-lookup"><span data-stu-id="ba40f-p103">Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.</span></span>

<span data-ttu-id="ba40f-228">Надстройки, отправляемые в AppSource, также должны включать элемент [SupportUrl](../reference/manifest/supporturl.md).</span><span class="sxs-lookup"><span data-stu-id="ba40f-228">Add-ins submitted to AppSource must also include the [SupportUrl](../reference/manifest/supporturl.md) element.</span></span> <span data-ttu-id="ba40f-229">Дополнительные сведения см. в статье [Политики проверки для приложений и надстроек, отправляемых в AppSource](/legal/marketplace/certification-policies).</span><span class="sxs-lookup"><span data-stu-id="ba40f-229">For more information, see [Validation policies for apps and add-ins submitted to AppSource](/legal/marketplace/certification-policies).</span></span>

<span data-ttu-id="ba40f-230">Чтобы указать домены, отличные от указанного в элементе [SourceLocation](../reference/manifest/sourcelocation.md) для сценариев проверки подлинности, используйте только элемент [AppDomains](../reference/manifest/appdomains.md).</span><span class="sxs-lookup"><span data-stu-id="ba40f-230">Only use the [AppDomains](../reference/manifest/appdomains.md) element to specify domains other than the one specified in the [SourceLocation](../reference/manifest/sourcelocation.md) element for authentication scenarios.</span></span>

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a><span data-ttu-id="ba40f-231">Укажите домены, которые необходимо открыть в окне надстройки</span><span class="sxs-lookup"><span data-stu-id="ba40f-231">Specify domains you want to open in the add-in window</span></span>

<span data-ttu-id="ba40f-232">В Office в Интернете область задач может открывать любой URL-адрес.</span><span class="sxs-lookup"><span data-stu-id="ba40f-232">When running in Office on the web, your task pane can be navigated to any URL.</span></span> <span data-ttu-id="ba40f-233">Если на платформах для настольных компьютеров надстройка пытается перейти на URL-адрес в домене, отличном от домена, где размещена начальная страница (указан в элементе [SourceLocation](../reference/manifest/sourcelocation.md) файла манифеста), этот URL-адрес откроется в новом окне браузера, а не в области надстроек приложения Office.</span><span class="sxs-lookup"><span data-stu-id="ba40f-233">However, in desktop platforms, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](../reference/manifest/sourcelocation.md) element of the manifest file), that URL opens in a new browser window outside the add-in pane of the Office application.</span></span>

<span data-ttu-id="ba40f-234">Чтобы переопределить это поведение, укажите все домены, которые должны открываться в окне надстройки, в списке доменов в элементе [AppDomains](../reference/manifest/appdomains.md) файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="ba40f-234">To override this (desktop Office) behavior, specify each domain you want to open in the add-in window in the list of domains specified in the [AppDomains](../reference/manifest/appdomains.md) element of the manifest file.</span></span> <span data-ttu-id="ba40f-235">URL-адреса в доменах из списка будут открываться в области задач как в классическом Office, так и в Office в Интернете.</span><span class="sxs-lookup"><span data-stu-id="ba40f-235">If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both Office on the web and desktop.</span></span> <span data-ttu-id="ba40f-236">URL-адреса в доменах не из списка будут открываться в новом окне браузера (не в области надстроек) в классическом Office.</span><span class="sxs-lookup"><span data-stu-id="ba40f-236">If it tries to go to a URL that isn't in the list, then, in desktop Office, that URL opens in a new browser window (outside the add-in pane).</span></span>

> [!NOTE]
> <span data-ttu-id="ba40f-237">Из этого правила есть два исключения:</span><span class="sxs-lookup"><span data-stu-id="ba40f-237">There are two exceptions to this behavior:</span></span>
>
> - <span data-ttu-id="ba40f-238">Это относится только к корневой области надстройки.</span><span class="sxs-lookup"><span data-stu-id="ba40f-238">It applies only to the root pane of the add-in.</span></span> <span data-ttu-id="ba40f-239">Если в страницу надстройки внедрен iframe, его можно перенаправить на любой URL-адрес, независимо от того, указан ли он в элементе **AppDomains**, даже в классической версии Office.</span><span class="sxs-lookup"><span data-stu-id="ba40f-239">If there is an iframe embedded in the add-in page, the iframe can be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>
> - <span data-ttu-id="ba40f-240">Если диалоговое окно открыто с помощью API [displayDialogAsync](/javascript/api/office/office.ui?view=common-js#displaydialogasync-startaddress--options--callback-), URL-адрес, передаваемый методу, должен находиться в том же домене, что и надстройка. Затем диалоговое окно можно перенаправить на любой URL-адрес, независимо от того, указан ли он в элементе **AppDomains**, даже в классической версии Office.</span><span class="sxs-lookup"><span data-stu-id="ba40f-240">When a dialog is opened with the [displayDialogAsync](/javascript/api/office/office.ui?view=common-js#displaydialogasync-startaddress--options--callback-) API, the URL that is passed to the method must be in the same domain as the add-in, but the dialog can then be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>

<span data-ttu-id="ba40f-241">В приведенном ниже примере XML-манифеста главная страница надстройки размещена в домене `https://www.contoso.com`, указанном в элементе **SourceLocation**.</span><span class="sxs-lookup"><span data-stu-id="ba40f-241">The following XML manifest example hosts its main add-in page in the `https://www.contoso.com` domain as specified in the **SourceLocation** element.</span></span> <span data-ttu-id="ba40f-242">В нем также указан домен `https://www.northwindtraders.com` с помощью элемента [AppDomain](../reference/manifest/appdomain.md) из списка **AppDomains**.</span><span class="sxs-lookup"><span data-stu-id="ba40f-242">It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](../reference/manifest/appdomain.md) element within the **AppDomains** element list.</span></span> <span data-ttu-id="ba40f-243">Страница в домене `www.northwindtraders.com` будет открываться в области надстроек даже в классической версии Office.</span><span class="sxs-lookup"><span data-stu-id="ba40f-243">If the add-in goes to a page in the `www.northwindtraders.com` domain, that page opens in the add-in pane, even in Office desktop.</span></span>

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
  <Id>c6890c26-5bbb-40ed-a321-37f07909a2f0</Id>
  <Version>1.0</Version>
  <ProviderName>Contoso, Ltd</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Northwind Traders Excel" />
  <Description DefaultValue="Search Northwind Traders data from Excel"/>
  <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
  <AppDomains>
    <AppDomain>https://www.northwindtraders.com</AppDomain>
  </AppDomains>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://www.contoso.com/search_app/Default.aspx" />
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

## <a name="specify-domains-from-which-officejs-api-calls-are-made"></a><span data-ttu-id="ba40f-244">Указание доменов, из которых выполняются вызовы API Office.js</span><span class="sxs-lookup"><span data-stu-id="ba40f-244">Specify domains from which Office.js API calls are made</span></span>

<span data-ttu-id="ba40f-245">Ваша надстройка может выполнять вызовы API Office.js из домена, указанного в элементе [SourceLocation](../reference/manifest/sourcelocation.md) файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="ba40f-245">Your add-in can make Office.js API calls from the domain referenced in the [SourceLocation](../reference/manifest/sourcelocation.md) element of the manifest file.</span></span> <span data-ttu-id="ba40f-246">Если в вашей надстройке есть другие блоки IFrame, которым требуется доступ к API Office.js, добавьте домен этого исходного URL-адреса в список, указанный в элементе [AppDomains](../reference/manifest/appdomains.md) файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="ba40f-246">If you have other IFrames within your add-in that need to access Office.js APIs, add the domain of that source URL to the list specified in the [AppDomains](../reference/manifest/appdomains.md) element of the manifest file.</span></span> <span data-ttu-id="ba40f-247">Если блок IFrame с источником, не содержащимся в списке `AppDomains`, попытается выполнить вызов API Office.js, надстройка получит [ошибку об отказе в разрешении](../reference/javascript-api-for-office-error-codes.md).</span><span class="sxs-lookup"><span data-stu-id="ba40f-247">If an IFrame with a source not contained in the `AppDomains` list attempts to make an Office.js API call, then the add-in will receive a [permission denied error](../reference/javascript-api-for-office-error-codes.md).</span></span>

## <a name="manifest-v11-xml-file-examples-and-schemas"></a><span data-ttu-id="ba40f-248">XML-файлы манифеста версии 1.1: примеры и схемы</span><span class="sxs-lookup"><span data-stu-id="ba40f-248">Manifest v1.1 XML file examples and schemas</span></span>

<span data-ttu-id="ba40f-249">Ниже показаны примеры XML-файлов манифеста версии 1.1 для надстроек области задач, контентных надстроек и надстроек Outlook.</span><span class="sxs-lookup"><span data-stu-id="ba40f-249">The following sections show examples of manifest v1.1 XML files for content, task pane, and Outlook add-ins.</span></span>

# <a name="task-pane"></a>[<span data-ttu-id="ba40f-250">Области задач</span><span class="sxs-lookup"><span data-stu-id="ba40f-250">Task pane</span></span>](#tab/tabid-1)

[<span data-ttu-id="ba40f-251">Схемы манифестов надстроек</span><span class="sxs-lookup"><span data-stu-id="ba40f-251">Add-in manifest schemas</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">

  <!-- See https://github.com/OfficeDev/Office-Add-in-Commands-Samples for documentation-->

  <!-- BeginBasicSettings: Add-in metadata, used for all versions of Office unless override provided -->

  <!--IMPORTANT! Id must be unique for your add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
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

          <!--PrimaryCommandSurface==Main Office app ribbon-->
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
                  <!--Icons. Required sizes: 16, 32, 80; optional: 20, 24, 40, 48, 64. You should provide as many sizes as possible for a great user experience. -->
                  <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                  <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                  <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
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
                    <bt:Image size="16" resid="Contoso.FunctionButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.FunctionButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.FunctionButton.Icon80" />
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
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
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
                    <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                    <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                    <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
                  </Icon>
                  <Items>
                    <Item id="Contoso.Menu.Item1">
                      <Label resid="Contoso.Item1.Label"/>
                      <Supertip>
                        <Title resid="Contoso.Item1.Label" />
                        <Description resid="Contoso.Item1.Tooltip" />
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
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
                        <bt:Image size="16" resid="Contoso.TaskpaneButton.Icon16" />
                        <bt:Image size="32" resid="Contoso.TaskpaneButton.Icon32" />
                        <bt:Image size="80" resid="Contoso.TaskpaneButton.Icon80" />
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
              <!-- If validating with XSD it needs to be at the end -->
              <Label resid="Contoso.Tab1.TabLabel" />
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Contoso.TaskpaneButton.Icon16" DefaultValue="https://myCDN/Images/Button16x16.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon32" DefaultValue="https://myCDN/Images/Button32x32.png" />
        <bt:Image id="Contoso.TaskpaneButton.Icon80" DefaultValue="https://myCDN/Images/Button80x80.png" />
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

# <a name="content"></a>[<span data-ttu-id="ba40f-252">Контент</span><span class="sxs-lookup"><span data-stu-id="ba40f-252">Content</span></span>](#tab/tabid-2)

[<span data-ttu-id="ba40f-253">Схемы манифестов надстроек</span><span class="sxs-lookup"><span data-stu-id="ba40f-253">Add-in manifest schemas</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="ContentApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
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

# <a name="mail"></a><span data-ttu-id="ba40f-254">[почта](#tab/tabid-3);</span><span class="sxs-lookup"><span data-stu-id="ba40f-254">[Mail](#tab/tabid-3)</span></span>

[<span data-ttu-id="ba40f-255">Схемы манифестов надстроек</span><span class="sxs-lookup"><span data-stu-id="ba40f-255">Add-in manifest schemas</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns=
  "http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xsi:type="MailApp">
  <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
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
      directement depuis Outlook."/>
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

## <a name="validate-an-office-add-ins-manifest"></a><span data-ttu-id="ba40f-256">Проверка манифеста надстройки Office</span><span class="sxs-lookup"><span data-stu-id="ba40f-256">Validate an Office Add-in's manifest</span></span>

<span data-ttu-id="ba40f-257">Сведения о проверке манифеста с помощью [определения схемы XML (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8) см. в статье [Проверка манифеста надстройки Office](../testing/troubleshoot-manifest.md).</span><span class="sxs-lookup"><span data-stu-id="ba40f-257">For information about validating a manifest against the [XML Schema Definition (XSD)](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8), see [Validate an Office Add-in's manifest](../testing/troubleshoot-manifest.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="ba40f-258">См. также</span><span class="sxs-lookup"><span data-stu-id="ba40f-258">See also</span></span>

* [<span data-ttu-id="ba40f-259">Определение правильного порядка элементов манифеста</span><span class="sxs-lookup"><span data-stu-id="ba40f-259">How to find the proper order of manifest elements</span></span>](manifest-element-ordering.md)
* <span data-ttu-id="ba40f-260">[Создание команд надстройки в манифесте][команды надстройки]</span><span class="sxs-lookup"><span data-stu-id="ba40f-260">[Create add-in commands in your manifest][add-in commands]</span></span>
* [<span data-ttu-id="ba40f-261">Указание приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="ba40f-261">Specify Office applications and API requirements</span></span>](specify-office-hosts-and-api-requirements.md)
* [<span data-ttu-id="ba40f-262">Локализация надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="ba40f-262">Localization for Office Add-ins</span></span>](localization.md)
* [<span data-ttu-id="ba40f-263">Справочная схема по манифестам надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="ba40f-263">Schema reference for Office Add-ins manifests</span></span>](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8)
* [<span data-ttu-id="ba40f-264">Обновление API и версии манифеста</span><span class="sxs-lookup"><span data-stu-id="ba40f-264">Update API and manifest version</span></span>](update-your-javascript-api-for-office-and-manifest-schema-version.md)
* [<span data-ttu-id="ba40f-265">Определение аналогичной надстройки COM</span><span class="sxs-lookup"><span data-stu-id="ba40f-265">Identify an equivalent COM add-in</span></span>](make-office-add-in-compatible-with-existing-com-add-in.md)
* [<span data-ttu-id="ba40f-266">Запрос разрешений на использование API в надстройках</span><span class="sxs-lookup"><span data-stu-id="ba40f-266">Requesting permissions for API use in add-ins</span></span>](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)
* [<span data-ttu-id="ba40f-267">Проверка манифеста надстройки Office</span><span class="sxs-lookup"><span data-stu-id="ba40f-267">Validate an Office Add-in's manifest</span></span>](../testing/troubleshoot-manifest.md)

[команды надстройки]: create-addin-commands.md
[add-in commands]: create-addin-commands.md
