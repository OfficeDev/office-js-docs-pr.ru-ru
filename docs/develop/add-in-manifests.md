---
title: XML-манифест надстроек Office
description: ''
ms.date: 02/09/2018
ms.openlocfilehash: 449c6ae3f98383ddff5f866cc47d19ea82ea89c2
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925404"
---
# <a name="office-add-ins-xml-manifest"></a><span data-ttu-id="44452-102">XML-манифест надстроек Office</span><span class="sxs-lookup"><span data-stu-id="44452-102">Office Add-ins XML manifest</span></span>

<span data-ttu-id="44452-103">XML-файл манифеста надстройки Office описывает способ ее активации, когда пользователь устанавливает и использует эту надстройку для работы с документами и приложениями Office.</span><span class="sxs-lookup"><span data-stu-id="44452-103">The XML manifest file of an Office Add-in describes how your add-in should be activated when an end user installs and uses it with Office documents and applications.</span></span>

<span data-ttu-id="44452-104">С помощью такого XML-файла манифеста надстройка Office может выполнять следующие действия:</span><span class="sxs-lookup"><span data-stu-id="44452-104">An XML manifest file based on this schema enables an Office Add-in to do the following:</span></span>

* <span data-ttu-id="44452-105">Предоставлять идентификатор, версию, описание, отображаемое имя и языковой стандарт по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="44452-105">Describe itself by providing an ID, version, description, display name, and default locale.</span></span>

* <span data-ttu-id="44452-106">Указывать изображения, используемые для фирменного оформления надстройки, и значки, используемые для [команд надстройки][] в ленте Office.</span><span class="sxs-lookup"><span data-stu-id="44452-106">Specify the images used for branding the Add-in and iconography used for [Add-in Commands][] in the Office Ribbon.</span></span>

* <span data-ttu-id="44452-107">Указывать, как надстройка интегрируется с Office, включая создаваемые ею элементы пользовательского интерфейса, например кнопки на ленте.</span><span class="sxs-lookup"><span data-stu-id="44452-107">Specify how the add-in integrates with Office, including any custom UI, such as ribbon buttons the add-in creates.</span></span>

* <span data-ttu-id="44452-108">Определять запрошенные размеры по умолчанию для контентных надстроек, а также запрошенную высоту для надстроек Outlook.</span><span class="sxs-lookup"><span data-stu-id="44452-108">Specify the requested default dimensions for content add-ins, and requested height for Outlook add-ins.</span></span>

* <span data-ttu-id="44452-109">Объявлять разрешения, в которых нуждается Надстройка Office, например чтение или запись документа.</span><span class="sxs-lookup"><span data-stu-id="44452-109">Declare permissions that the Office Add-in requires, such as reading or writing to the document.</span></span>

* <span data-ttu-id="44452-110">В случае надстроек Outlook необходимо определить одно или несколько правил, указывающих контекст, в котором эти надстройки будут активироваться и взаимодействовать с сообщением, сведениями о встрече или приглашением на собрание.</span><span class="sxs-lookup"><span data-stu-id="44452-110">For Outlook add-ins, define the rule or rules that specify the context in which they will be activated and interact with a message, appointment, or meeting request item.</span></span>

> [!NOTE]
> <span data-ttu-id="44452-p101">Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource и сделать ее доступной в интерфейсе Office, убедитесь, что она соответствует [политикам проверки AppSource](https://docs.microsoft.com/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и на [странице со сведениями о доступности и ведущих приложениях для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="44452-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](https://docs.microsoft.com/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="required-elements"></a><span data-ttu-id="44452-113">Обязательные элементы</span><span class="sxs-lookup"><span data-stu-id="44452-113">Required elements</span></span>

<span data-ttu-id="44452-114">В приведенной ниже таблице указаны обязательные элементы для трех типов надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="44452-114">The following table specifies the elements that are required for the three types of Office Add-ins.</span></span>

### <a name="required-elements-by-office-add-in-type"></a><span data-ttu-id="44452-115">Обязательные элементы по типам надстроек Office</span><span class="sxs-lookup"><span data-stu-id="44452-115">Required elements by Office Add-in type</span></span>

| <span data-ttu-id="44452-116">Элемент</span><span class="sxs-lookup"><span data-stu-id="44452-116">Element</span></span>                                                                                      | <span data-ttu-id="44452-117">Контент</span><span class="sxs-lookup"><span data-stu-id="44452-117">Content</span></span> | <span data-ttu-id="44452-118">Область задач</span><span class="sxs-lookup"><span data-stu-id="44452-118">Task pane</span></span> | <span data-ttu-id="44452-119">Outlook</span><span class="sxs-lookup"><span data-stu-id="44452-119">Outlook</span></span> |
| :------------------------------------------------------------------------------------------- | :-----: | :-------: | :-----: |
| <span data-ttu-id="44452-120">[OfficeApp][]</span><span class="sxs-lookup"><span data-stu-id="44452-120">[OfficeApp][]</span></span>                                                                                |    <span data-ttu-id="44452-121">X</span><span class="sxs-lookup"><span data-stu-id="44452-121">X</span></span>    |     <span data-ttu-id="44452-122">X</span><span class="sxs-lookup"><span data-stu-id="44452-122">X</span></span>     |    <span data-ttu-id="44452-123">X</span><span class="sxs-lookup"><span data-stu-id="44452-123">X</span></span>    |
| <span data-ttu-id="44452-124">[Id][]</span><span class="sxs-lookup"><span data-stu-id="44452-124">[Id][]</span></span>                                                                                       |    <span data-ttu-id="44452-125">X</span><span class="sxs-lookup"><span data-stu-id="44452-125">X</span></span>    |     <span data-ttu-id="44452-126">X</span><span class="sxs-lookup"><span data-stu-id="44452-126">X</span></span>     |    <span data-ttu-id="44452-127">X</span><span class="sxs-lookup"><span data-stu-id="44452-127">X</span></span>    |
| <span data-ttu-id="44452-128">[Версия][]</span><span class="sxs-lookup"><span data-stu-id="44452-128">[Version][]</span></span>                                                                                  |    <span data-ttu-id="44452-129">X</span><span class="sxs-lookup"><span data-stu-id="44452-129">X</span></span>    |     <span data-ttu-id="44452-130">X</span><span class="sxs-lookup"><span data-stu-id="44452-130">X</span></span>     |    <span data-ttu-id="44452-131">X</span><span class="sxs-lookup"><span data-stu-id="44452-131">X</span></span>    |
| <span data-ttu-id="44452-132">[ProviderName][]</span><span class="sxs-lookup"><span data-stu-id="44452-132">[ProviderName][]</span></span>                                                                             |    <span data-ttu-id="44452-133">X</span><span class="sxs-lookup"><span data-stu-id="44452-133">X</span></span>    |     <span data-ttu-id="44452-134">X</span><span class="sxs-lookup"><span data-stu-id="44452-134">X</span></span>     |    <span data-ttu-id="44452-135">X</span><span class="sxs-lookup"><span data-stu-id="44452-135">X</span></span>    |
| <span data-ttu-id="44452-136">[DefaultLocale][]</span><span class="sxs-lookup"><span data-stu-id="44452-136">[DefaultLocale][]</span></span>                                                                            |    <span data-ttu-id="44452-137">X</span><span class="sxs-lookup"><span data-stu-id="44452-137">X</span></span>    |     <span data-ttu-id="44452-138">X</span><span class="sxs-lookup"><span data-stu-id="44452-138">X</span></span>     |    <span data-ttu-id="44452-139">X</span><span class="sxs-lookup"><span data-stu-id="44452-139">X</span></span>    |
| <span data-ttu-id="44452-140">[DisplayName][]</span><span class="sxs-lookup"><span data-stu-id="44452-140">[DisplayName][]</span></span>                                                                              |    <span data-ttu-id="44452-141">X</span><span class="sxs-lookup"><span data-stu-id="44452-141">X</span></span>    |     <span data-ttu-id="44452-142">X</span><span class="sxs-lookup"><span data-stu-id="44452-142">X</span></span>     |    <span data-ttu-id="44452-143">X</span><span class="sxs-lookup"><span data-stu-id="44452-143">X</span></span>    |
| <span data-ttu-id="44452-144">[Описание][]</span><span class="sxs-lookup"><span data-stu-id="44452-144">[Description][]</span></span>                                                                              |    <span data-ttu-id="44452-145">X</span><span class="sxs-lookup"><span data-stu-id="44452-145">X</span></span>    |     <span data-ttu-id="44452-146">X</span><span class="sxs-lookup"><span data-stu-id="44452-146">X</span></span>     |    <span data-ttu-id="44452-147">X</span><span class="sxs-lookup"><span data-stu-id="44452-147">X</span></span>    |
| <span data-ttu-id="44452-148">[IconUrl][]</span><span class="sxs-lookup"><span data-stu-id="44452-148">[IconUrl][]</span></span>                                                                                  |    <span data-ttu-id="44452-149">X</span><span class="sxs-lookup"><span data-stu-id="44452-149">X</span></span>    |     <span data-ttu-id="44452-150">X</span><span class="sxs-lookup"><span data-stu-id="44452-150">X</span></span>     |    <span data-ttu-id="44452-151">X</span><span class="sxs-lookup"><span data-stu-id="44452-151">X</span></span>    |
| <span data-ttu-id="44452-152">[HighResolutionIconUrl][]</span><span class="sxs-lookup"><span data-stu-id="44452-152">[HighResolutionIconUrl][]</span></span>                                                                    |    <span data-ttu-id="44452-153">X</span><span class="sxs-lookup"><span data-stu-id="44452-153">X</span></span>    |     <span data-ttu-id="44452-154">X</span><span class="sxs-lookup"><span data-stu-id="44452-154">X</span></span>     |    <span data-ttu-id="44452-155">X</span><span class="sxs-lookup"><span data-stu-id="44452-155">X</span></span>    |
| <span data-ttu-id="44452-156">[DefaultSettings (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="44452-156">[DefaultSettings (ContentApp)][]</span></span><br/><span data-ttu-id="44452-157">[DefaultSettings (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="44452-157">[DefaultSettings (TaskPaneApp)][]</span></span>                       |    <span data-ttu-id="44452-158">X</span><span class="sxs-lookup"><span data-stu-id="44452-158">X</span></span>    |     <span data-ttu-id="44452-159">X</span><span class="sxs-lookup"><span data-stu-id="44452-159">X</span></span>     |         |
| <span data-ttu-id="44452-160">[SourceLocation (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="44452-160">[SourceLocation (ContentApp)][]</span></span><br/><span data-ttu-id="44452-161">[SourceLocation (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="44452-161">[SourceLocation (TaskPaneApp)][]</span></span>                         |    <span data-ttu-id="44452-162">X</span><span class="sxs-lookup"><span data-stu-id="44452-162">X</span></span>    |     <span data-ttu-id="44452-163">X</span><span class="sxs-lookup"><span data-stu-id="44452-163">X</span></span>     |         |
| <span data-ttu-id="44452-164">[DesktopSettings][]</span><span class="sxs-lookup"><span data-stu-id="44452-164">[DesktopSettings][]</span></span>                                                                          |         |           |    <span data-ttu-id="44452-165">X</span><span class="sxs-lookup"><span data-stu-id="44452-165">X</span></span>    |
| <span data-ttu-id="44452-166">[SourceLocation (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="44452-166">[SourceLocation (MailApp)][]</span></span>                                                                 |         |           |    <span data-ttu-id="44452-167">X</span><span class="sxs-lookup"><span data-stu-id="44452-167">X</span></span>    |
| <span data-ttu-id="44452-168">[Разрешения (ContentApp)][]</span><span class="sxs-lookup"><span data-stu-id="44452-168">[Permissions (ContentApp)][]</span></span><br/><span data-ttu-id="44452-169">[Разрешения (TaskPaneApp)][]</span><span class="sxs-lookup"><span data-stu-id="44452-169">[Permissions (TaskPaneApp)][]</span></span><br/><span data-ttu-id="44452-170">[Разрешения (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="44452-170">[Permissions (MailApp)][]</span></span> |    <span data-ttu-id="44452-171">X</span><span class="sxs-lookup"><span data-stu-id="44452-171">X</span></span>    |     <span data-ttu-id="44452-172">X</span><span class="sxs-lookup"><span data-stu-id="44452-172">X</span></span>     |    <span data-ttu-id="44452-173">X</span><span class="sxs-lookup"><span data-stu-id="44452-173">X</span></span>    |
| <span data-ttu-id="44452-174">[Правило (RuleCollection)][]</span><span class="sxs-lookup"><span data-stu-id="44452-174">[Rule (RuleCollection)][]</span></span><br/><span data-ttu-id="44452-175">[Правило (MailApp)][]</span><span class="sxs-lookup"><span data-stu-id="44452-175">[Rule (MailApp)][]</span></span>                                             |         |           |    <span data-ttu-id="44452-176">X</span><span class="sxs-lookup"><span data-stu-id="44452-176">X</span></span>    |
| <span data-ttu-id="44452-177">[Требования (MailApp)\*][]</span><span class="sxs-lookup"><span data-stu-id="44452-177">[Requirements (MailApp)\*][]</span></span>                                                                  |         |           |    <span data-ttu-id="44452-178">X</span><span class="sxs-lookup"><span data-stu-id="44452-178">X</span></span>    |
| <span data-ttu-id="44452-179">[Установка\*][]</span><span class="sxs-lookup"><span data-stu-id="44452-179">[Set\*][]</span></span><br/><span data-ttu-id="44452-180">[Установки (MailAppRequirements)\*][]</span><span class="sxs-lookup"><span data-stu-id="44452-180">[Sets (MailAppRequirements)\*][]</span></span>                                                 |         |           |    <span data-ttu-id="44452-181">X</span><span class="sxs-lookup"><span data-stu-id="44452-181">X</span></span>    |
| <span data-ttu-id="44452-182">[Форма\*][]</span><span class="sxs-lookup"><span data-stu-id="44452-182">[Form\*][]</span></span><br/><span data-ttu-id="44452-183">[FormSettings\*][]</span><span class="sxs-lookup"><span data-stu-id="44452-183">[formsettings\*][]</span></span>                                                              |         |           |    <span data-ttu-id="44452-184">X</span><span class="sxs-lookup"><span data-stu-id="44452-184">X</span></span>    |
| <span data-ttu-id="44452-185">[Установка (требований)\*][]</span><span class="sxs-lookup"><span data-stu-id="44452-185">[Sets (Requirements)\*][]</span></span>                                                                     |    <span data-ttu-id="44452-186">X</span><span class="sxs-lookup"><span data-stu-id="44452-186">X</span></span>    |     <span data-ttu-id="44452-187">X</span><span class="sxs-lookup"><span data-stu-id="44452-187">X</span></span>     |         |
| <span data-ttu-id="44452-188">[Хосты\*][]</span><span class="sxs-lookup"><span data-stu-id="44452-188">[Hosts\*][]</span></span>                                                                                   |    <span data-ttu-id="44452-189">X</span><span class="sxs-lookup"><span data-stu-id="44452-189">X</span></span>    |     <span data-ttu-id="44452-190">X</span><span class="sxs-lookup"><span data-stu-id="44452-190">X</span></span>     |         |

<span data-ttu-id="44452-191">_\*Элемент добавлен в схеме манифеста для надстроек Office версии 1.1._</span><span class="sxs-lookup"><span data-stu-id="44452-191">_\*Added in the Office Add-in Manifest Schema version 1.1._</span></span>

<!-- Links for above table -->

[officeapp]: https://dev.office.com/reference/add-ins/manifest/officeapp
[идентификатор id]: https://dev.office.com/reference/add-ins/manifest/id
[id]: https://dev.office.com/reference/add-ins/manifest/id
[версия]: https://dev.office.com/reference/add-ins/manifest/version
[version]: https://dev.office.com/reference/add-ins/manifest/version
[providername]: https://dev.office.com/reference/add-ins/manifest/providername
[defaultlocale]: https://dev.office.com/reference/add-ins/manifest/defaultlocale
[displayname]: https://dev.office.com/reference/add-ins/manifest/displayname
[описание]: https://dev.office.com/reference/add-ins/manifest/description
[description]: https://dev.office.com/reference/add-ins/manifest/description
[iconurl]: https://dev.office.com/reference/add-ins/manifest/iconurl
[highresolutioniconurl]: https://dev.office.com/reference/add-ins/manifest/highresolutioniconurl
[defaultsettings (contentapp)]: https://dev.office.com/reference/add-ins/manifest/defaultsettings
[defaultsettings (taskpaneapp)]: https://dev.office.com/reference/add-ins/manifest/defaultsettings
[sourcelocation (contentapp)]: https://dev.office.com/reference/add-ins/manifest/sourcelocation
[sourcelocation (taskpaneapp)]: https://dev.office.com/reference/add-ins/manifest/sourcelocation
[desktopsettings]: https://msdn.microsoft.com/library/da9fd085-b8cc-2be0-d329-2aa1ef5d3f1c(Office.15).aspx
[sourcelocation (mailapp)]: http://msdn.microsoft.com/library/3792d389-bebd-d19a-9d90-35b7a0bfc623%28Office.15%29.aspx
[разрешения (contentapp)]: https://dev.office.com/reference/add-ins/manifest/permissions
[permissions (contentapp)]: https://dev.office.com/reference/add-ins/manifest/permissions
[разрешения (taskpaneapp)]: https://dev.office.com/reference/add-ins/manifest/permissions
[permissions (taskpaneapp)]: https://dev.office.com/reference/add-ins/manifest/permissions
[разрешения (mailapp)]: https://dev.office.com/reference/add-ins/manifest/permissions
[permissions (mailapp)]: https://dev.office.com/reference/add-ins/manifest/permissions
[правило (rulecollection)]: https://dev.office.com/reference/add-ins/manifest/rule
[rule (rulecollection)]: https://dev.office.com/reference/add-ins/manifest/rule
[правило  (mailapp)]: https://dev.office.com/reference/add-ins/manifest/rule
[rule (mailapp)]: https://dev.office.com/reference/add-ins/manifest/rule
[требования (mailapp)\*]: https://dev.office.com/reference/add-ins/manifest/requirements
[requirements (mailapp)\*]: https://dev.office.com/reference/add-ins/manifest/requirements
[установка\*]: https://dev.office.com/reference/add-ins/manifest/set
[set\*]: https://dev.office.com/reference/add-ins/manifest/set
[установки (mailapprequirements)\*]: https://dev.office.com/reference/add-ins/manifest/sets
[sets (mailapprequirements)\*]: https://dev.office.com/reference/add-ins/manifest/sets
[форма\*]: https://dev.office.com/reference/add-ins/manifest/form
[form\*]: https://dev.office.com/reference/add-ins/manifest/form
[formsettings*]: https://dev.office.com/reference/add-ins/manifest/formsettings
[Установка (требований)\*]: https://dev.office.com/reference/add-ins/manifest/sets
[sets (requirements)\*]: https://dev.office.com/reference/add-ins/manifest/sets
[хосты\*]: https://dev.office.com/reference/add-ins/manifest/hosts
[hosts\*]: https://dev.office.com/reference/add-ins/manifest/hosts

## <a name="hosting-requirements"></a><span data-ttu-id="44452-219">Требования к размещению</span><span class="sxs-lookup"><span data-stu-id="44452-219">Hosting requirements</span></span>

<span data-ttu-id="44452-220">Все URI изображений, в частности используемых для [команд надстройки][], должны поддерживать кэширование.</span><span class="sxs-lookup"><span data-stu-id="44452-220">All image URIs, such as those used for [Add-in Commands][], must support caching.</span></span> <span data-ttu-id="44452-221">Сервер с изображением не должен возвращать заголовок `Cache-Control`, содержащий `no-cache`, `no-store` или подобные параметры в HTTP-отклике.</span><span class="sxs-lookup"><span data-stu-id="44452-221">The server hosting the image should not return a `Cache-Control` header specifying `no-cache`, `no-store`, or similar options in the HTTP response.</span></span>

<span data-ttu-id="44452-222">Все URL-адреса, например адреса исходных файлов, указанные в элементе [SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation), должны быть **защищены с помощью SSL (HTTPS)**.</span><span class="sxs-lookup"><span data-stu-id="44452-222">All URLs, such as the source file locations specified in the [SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation) element, should be **SSL-secured (HTTPS)**.</span></span> [!include[HTTPS guidance](../includes/https-guidance.md)]

## <a name="best-practices-for-submitting-to-appsource"></a><span data-ttu-id="44452-223">Рекомендации по отправке решений в AppSource</span><span class="sxs-lookup"><span data-stu-id="44452-223">Best practices for submitting to AppSource</span></span>

<span data-ttu-id="44452-p103">Убедитесь, что идентификатор надстройки представляет собой допустимый и уникальный GUID. В Интернете доступно множество генераторов, с помощью которых можно создать уникальный GUID.</span><span class="sxs-lookup"><span data-stu-id="44452-p103">Make sure that the add-in ID is a valid and unique GUID. Various GUID generator tools are available on the web that you can use to create a unique GUID.</span></span>

<span data-ttu-id="44452-226">Надстройки, отправляемые в AppSource, также должны включать элемент [SupportUrl](https://dev.office.com/reference/add-ins/manifest/supporturl).</span><span class="sxs-lookup"><span data-stu-id="44452-226">Add-ins submitted to AppSource must also include the [SupportUrl](https://dev.office.com/reference/add-ins/manifest/supporturl) element.</span></span> <span data-ttu-id="44452-227">Дополнительные сведения см. в статье [Политики проверки для приложений и надстроек, отправляемых в AppSource](https://docs.microsoft.com/office/dev/store/validation-policies).</span><span class="sxs-lookup"><span data-stu-id="44452-227">For more information, see [Validation policies for apps and add-ins submitted to AppSource](https://docs.microsoft.com/office/dev/store/validation-policies).</span></span>

<span data-ttu-id="44452-228">Чтобы указать домены, отличные от указанного в элементе [SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation) для сценариев проверки подлинности, используйте только элемент [AppDomains](https://dev.office.com/reference/add-ins/manifest/appdomains).</span><span class="sxs-lookup"><span data-stu-id="44452-228">Only use the [AppDomains](https://dev.office.com/reference/add-ins/manifest/appdomains) element to specify domains other than the one specified in the [SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation) element for authentication scenarios.</span></span>

## <a name="specify-domains-you-want-to-open-in-the-add-in-window"></a><span data-ttu-id="44452-229">Указание доменов, которые необходимо открыть в окне надстройки</span><span class="sxs-lookup"><span data-stu-id="44452-229">Specify domains you want to open in the add-in window</span></span>

<span data-ttu-id="44452-230">При работе в Office Online ваша панель задач может быть перенесена на любой URL-адрес.</span><span class="sxs-lookup"><span data-stu-id="44452-230">When running in Office Online, your task pane can be navigated to any URL.</span></span> <span data-ttu-id="44452-231">Однако на классических платформах, если надстройка пытается перейти на URL-адрес в домене, отличном от домена, где размещена начальная страница (как указано в элементе [SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation) файла манифеста), этот URL-адрес откроется в новом окне веб-обозревателя, а не в панели надстроек ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="44452-231">By default, if your add-in tries to go to a URL in a domain other than the domain that hosts the start page (as specified in the [SourceLocation](https://dev.office.com/reference/add-ins/manifest/sourcelocation) element of the manifest file), that URL will open in a new browser window outside the add-in pane of the Office host application.</span></span>

<span data-ttu-id="44452-232">Чтобы переопределить эту реакцию на событие (в классической версии Office), добавьте каждый домен, который требуется открыть в окне надстройки, в список доменов в элементе [AppDomains](https://dev.office.com/reference/add-ins/manifest/appdomains) файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="44452-232">To override this behavior, specify each domain you want to open in the add-in window in the list of domains specified in the [AppDomains](https://dev.office.com/reference/add-ins/manifest/appdomains) element of the manifest file.</span></span> <span data-ttu-id="44452-233">Если надстройка пытается перейти по URL-адресу в домене, который находится в списке, она откроется в области задач как в классической версии Office, так и в Office Online.</span><span class="sxs-lookup"><span data-stu-id="44452-233">If the add-in tries to go to a URL in a domain that is in the list, then it opens in the task pane in both desktop Office and Office Online.</span></span> <span data-ttu-id="44452-234">Если она пытается перейти по URL-адресу, отсутствующему в списке, тогда в классической версии Office этот URL-адрес откроется в новом окне веб-обозревателя (вне панели надстройки).</span><span class="sxs-lookup"><span data-stu-id="44452-234">If the add-in tries to go to a URL in a domain that isn't in the list, that URL will open in a new browser window (outside the add-in pane).</span></span>

> [!NOTE]
> <span data-ttu-id="44452-235">Эта реакция на событие относится только к корневой панели надстройки.</span><span class="sxs-lookup"><span data-stu-id="44452-235">This behavior applies only to the root pane of the add-in.</span></span> <span data-ttu-id="44452-236">Если на странице надстройки есть элемент iframe, он может быть перенаправлен на любой URL-адрес независимо от того, указан ли он в **AppDomains** – даже в классической версии Office.</span><span class="sxs-lookup"><span data-stu-id="44452-236">If there is an iframe embedded in the add-in page, the iframe can be directed to any URL regardless of whether it is listed in **AppDomains**, even in desktop Office.</span></span>

<span data-ttu-id="44452-237">В приведенном ниже примере XML-манифеста главная страница надстройки размещена в домене `https://www.contoso.com`, как указано в элементе **SourceLocation**.</span><span class="sxs-lookup"><span data-stu-id="44452-237">The following XML manifest example hosts its main add-in page in the  `https://www.contoso.com` domain as specified in the **SourceLocation** element.</span></span> <span data-ttu-id="44452-238">В нем также указан домен `https://www.northwindtraders.com` с помощью элемента [AppDomain](https://dev.office.com/reference/add-ins/manifest/appdomain) из списка **AppDomains**.</span><span class="sxs-lookup"><span data-stu-id="44452-238">It also specifies the `https://www.northwindtraders.com` domain in an [AppDomain](https://dev.office.com/reference/add-ins/manifest/appdomain) element within the **AppDomains** element list.</span></span> <span data-ttu-id="44452-239">Если надстройка переходит на страницу в домене www.northwindtraders.com, эта страница открывается на панели надстройки – даже в классическом приложении Office.</span><span class="sxs-lookup"><span data-stu-id="44452-239">If the add-in goes to a page in the www.northwindtraders.com domain, that page will open in the add-in pane.</span></span>

```XML
<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
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

## <a name="manifest-v11-xml-file-examples-and-schemas"></a><span data-ttu-id="44452-240">Примеры и схемы XML-файла манифеста версии 1.1</span><span class="sxs-lookup"><span data-stu-id="44452-240">Manifest v1.1 XML file examples and schemas</span></span>
<span data-ttu-id="44452-241">Ниже показаны примеры XML-файлов манифеста версии 1.1 для надстроек области задач, контентных надстроек и надстроек Outlook.</span><span class="sxs-lookup"><span data-stu-id="44452-241">The following sections show examples of manifest v1.1 XML files for content, task pane, and Outlook add-ins.</span></span>

# <a name="task-panetabtabid-1"></a>[<span data-ttu-id="44452-242">Области задач</span><span class="sxs-lookup"><span data-stu-id="44452-242">Task pane</span></span>](#tab/tabid-1)

[<span data-ttu-id="44452-243">Схема манифеста приложения области задач</span><span class="sxs-lookup"><span data-stu-id="44452-243">Task pane app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/taskpane)

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
  <IconUrl DefaultValue="https://i.imgur.com/oZFS95h.png" />

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
        <!-- Form factor. Currenly only DesktopFormFactor is supported. We will add TabletFormFactor and PhoneFormFactor in the future-->
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
                  <!--This is what happens when the command is triggered (E.g. click on the Ribbon). Supported actions are ExecuteFuncion or ShowTaskpane-->
                  <!--Look at the FunctionFile.html page for reference on how to implement the function -->
                  - <Action xsi:type="ExecuteFunction">
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

# <a name="contenttabtabid-2"></a>[<span data-ttu-id="44452-244">Контентная</span><span class="sxs-lookup"><span data-stu-id="44452-244">Content</span></span>](#tab/tabid-2)

[<span data-ttu-id="44452-245">Схема манифеста контентного приложения</span><span class="sxs-lookup"><span data-stu-id="44452-245">Content app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/content)

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
  <IconUrl DefaultValue="https://contoso.com/ENUSIcon.png" />
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

# <a name="mailtabtabid-3"></a>[<span data-ttu-id="44452-246">Почтовая</span><span class="sxs-lookup"><span data-stu-id="44452-246">Mail</span></span>](#tab/tabid-3)

[<span data-ttu-id="44452-247">Схема манифеста почтового приложения</span><span class="sxs-lookup"><span data-stu-id="44452-247">Mail app manifest schema</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas/mail)

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
  <!-- Change the following line to specify    -->
  <!-- the web serverthat hosts the icon file. -->
  <IconUrl DefaultValue=
    "https://webserver/YouTube/YouTubeLogo.png"/>

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

## <a name="validate-and-troubleshoot-issues-with-your-manifest"></a><span data-ttu-id="44452-248">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="44452-248">Validate and troubleshoot issues with your manifest</span></span>

<span data-ttu-id="44452-p109">Сведения об устранении проблем, связанных с манифестом надстройки, см. в статье [Проверка манифеста и устранение связанных с ним неполадок](../testing/troubleshoot-manifest.md). Там указано, как проверить манифест согласно [XSD](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), а также как отладить манифест с помощью ведения журнала в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="44452-p109">For troubleshooting issues with your manifest, see [Validate and troubleshoot issues with your manifest](../testing/troubleshoot-manifest.md). There, you will find information on how to validate the manifest against the [XML Schema Definition (XSD)](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas), and also how to use runtime logging to debug the manifest.</span></span>

## <a name="see-also"></a><span data-ttu-id="44452-251">См. также</span><span class="sxs-lookup"><span data-stu-id="44452-251">See also</span></span>

* <span data-ttu-id="44452-252">[Создание команд надстройки в манифесте][команды надстройки]</span><span class="sxs-lookup"><span data-stu-id="44452-252">[Create add-in commands in your manifest][add-in commands]</span></span>
* [<span data-ttu-id="44452-253">Указание ведущих приложений Office и обязательных элементов API</span><span class="sxs-lookup"><span data-stu-id="44452-253">Specify Office hosts and API requirements</span></span>](specify-office-hosts-and-api-requirements.md)
* [<span data-ttu-id="44452-254">Локализация надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="44452-254">Localization for Office Add-ins</span></span>](localization.md)
* [<span data-ttu-id="44452-255">Справочная схема по манифестам надстроек для Office</span><span class="sxs-lookup"><span data-stu-id="44452-255">Schema reference for Office Add-ins manifests</span></span>](https://github.com/OfficeDev/office-js-docs-pr/tree/master/docs/overview/schemas)
* [<span data-ttu-id="44452-256">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="44452-256">Validate and troubleshoot issues with your manifest</span></span>](../testing/troubleshoot-manifest.md)

[команды надстройки]: create-addin-commands.md
[add-in commands]: create-addin-commands.md