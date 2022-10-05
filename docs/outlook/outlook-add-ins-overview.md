---
title: Обзор надстроек Outlook
description: Надстройки Outlook — это встраиваемые в Outlook решения от сторонних разработчиков, использующие нашу веб-платформу.
ms.date: 08/09/2022
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: fd17728f840188fbedfdeba7d3ee8f97852d702a
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467260"
---
# <a name="outlook-add-ins-overview"></a>Обзор надстроек Outlook

Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform. Outlook add-ins have three key aspects:

- В классических приложениях (Outlook для Windows и Mac), веб-приложениях (Microsoft 365 и Outlook.com) и мобильных решениях используются одинаковые логика надстроек и бизнес-логика.
- Надстройка Outlook состоит из манифеста, в котором описан способ интеграции надстройки с Outlook (например, при помощи кнопки или области задач), и кода JavaScript или HTML, который составляет пользовательский интерфейс и бизнес-логику надстройки.
- Пользователи и администраторы могут получать надстройки Outlook из [AppSource](https://appsource.microsoft.com) или [загружать их в неопубликованном виде](sideload-outlook-add-ins-for-testing.md).

Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows. Unlike COM add-ins, Outlook add-ins don't have any code physically installed on the user's device or Outlook client. For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, and then loads the JavaScript and HTML. The web components all run in the context of a browser in a sandbox.

The Outlook items that support add-ins include email messages, meeting requests, responses and cancellations, and appointments. Each Outlook add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a>Точки расширения

Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done.

- Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).

    **Надстройка с кнопками на ленте**

    ![Команда функции надстройки.](../images/uiless-command-shape.png)

- Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](contextual-outlook-add-ins.md).

    **Контекстная надстройка для выделенной сущности (адреса)**

    ![Показывает контекстное приложение на карте.](../images/outlook-detected-entity-card.png)

## <a name="mailbox-items-available-to-add-ins"></a>Элементы почтовых ящиков, доступные надстройкам

Надстройки Outlook активизируются при создании или чтении сообщения либо встречи, но не других типов элементов. При этом надстройки *не* активизируются, если текущий элемент сообщения в форме создания или просмотра имеет одну из следующих особенностей:

- Защищено службой управления правами на доступ к данным (IRM) или шифруется другими способами для защиты и доступа из Outlook на клиентах, отличных от Windows. Один из примеров — сообщение, подписанное цифровой подписью, так как в этом случае используется один из указанных выше механизмов.

[!INCLUDE [outlook-irm-add-in-activation](../includes/outlook-irm-add-in-activation.md)]

- Отчет или уведомление о доставке имеет класс сообщения IPM.Report.*, включая отчеты о доставке, о недоставке, а также уведомления о прочтении, о непрочтении и о задержке.

- MSG- или EML-файл, представляющий собой вложение в другое сообщение.

- MSG- или EML-файл, открытый из файловой системы.

- В [групповом почтовом ящике](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes), общем почтовом ящике\*, почтовом ящике другого пользователя\*, [архивном почтовом ящике](/office365/servicedescriptions/exchange-online-archiving-service-description/archive-client-and-compliance-&-security-feature-details?tabs=Archive-features#archive-mailbox) или общедоступной папке.

  > [!IMPORTANT]
  > \* Поддержка сценариев делегирования доступа (например, папок, полученных из почтового ящика другого пользователя) была представлена в [наборе требований 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8). Поддержка общих почтовых ящиков теперь доступна в предварительной версии в Outlook для Windows и Mac. Дополнительные сведения см. в статье ["Включение общих папок и сценариев общих почтовых ящиков"](delegate-access.md).

- Использование настраиваемой формы.

- Создано с помощью Simple MAPI. Simple MAPI используется, если пользователь Office создает или отправляет сообщение электронной почты из приложения Office в Windows, когда Outlook закрыт. Например, пользователь может создать сообщение электронной почты Outlook во время работы в Word, что запускает окно создания сообщения Outlook без запуска основного приложения Outlook. Однако если Outlook уже запущен, когда пользователь создает сообщение электронной почты из Word, это не сценарий Simple MAPI, поэтому надстройки Outlook работают в форме создания при условии, что выполнены другие требования к активации.

В общем случае Outlook может активировать надстройки в формах просмотра для элементов в папке "Отправленные", за исключением надстроек, активируемых на основании совпадений строк для известных сущностей. Дополнительные сведения о причинах этого см. в статье [Поддержка известных сущностей](match-strings-in-an-item-as-well-known-entities.md#support-for-well-known-entities).

В настоящее время при проектировании и внедрении надстроек для мобильных клиентов следует учитывать и другие факторы. Дополнительные сведения см. в статье ["Добавление поддержки мобильных устройств в надстройку Outlook"](add-mobile-support.md#compose-mode-and-appointments).

## <a name="supported-clients"></a>Поддерживаемые клиенты

Надстройки Outlook поддерживают Outlook 2013 или более поздней версии для Windows, Outlook 2016 или более поздней версии для Mac, Outlook в Интернете для локальной версии Exchange 2013 и более поздних версий, Outlook для iOS, Outlook для Android, Outlook в Интернете и Outlook.com. Не все новые функции поддерживаются сразу всеми [клиентами](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients). Просмотрите статьи и справочные материалы по API для этих функций, чтобы узнать, в каких приложениях они поддерживаются.

## <a name="get-started-building-outlook-add-ins"></a>Знакомство с разработкой надстроек Outlook

Чтобы приступить к разработке надстроек Outlook, попробуйте приведенные ниже ресурсы.

- [Краткое руководство](../quickstarts/outlook-quickstart.md) — создание простой надстройки области задач.
- [Учебник](../tutorials/outlook-tutorial.md) — узнайте, как создать надстройку, которая вставляет элементы gist с сайта GitHub в новое сообщение.

## <a name="see-also"></a>См. также

- [Сведения о программе для разработчиков Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
- [Рекомендации по разработке надстроек Office](../concepts/add-in-development-best-practices.md)
- [Рекомендации по проектированию надстроек Office](../design/add-in-design.md)
- [Лицензирование надстроек Office и SharePoint](/office/dev/store/license-your-add-ins)
- [Публикация надстройки Office](../publish/publish.md)
- [Публикация решений в AppSource и в Office](/office/dev/store/submit-to-the-office-store)
