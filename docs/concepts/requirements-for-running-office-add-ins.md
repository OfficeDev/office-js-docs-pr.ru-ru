---
title: Требования для запуска надстроек Office
description: Узнайте о требованиях к клиенту и серверу, которые необходимо выполнить Office надстройки.
ms.date: 02/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6e1bd7eb5f2949d6b0c70654c3aa3a276a3ee83c
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63742828"
---
# <a name="requirements-for-running-office-add-ins"></a>Требования для запуска надстроек Office

В этой статье описаны требования к программному обеспечению и устройствам для запуска надстроек Office.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

Для представления на высоком уровне о том, Office в настоящее время поддерживаются надстройки, см. в Office клиентского приложения и доступности платформы для [Office надстройки](../overview/office-add-in-availability.md).

## <a name="server-requirements"></a>Требования к серверу

Чтобы иметь возможность установить и запустить любую Надстройка Office, необходимо сначала развернуть файлы манифеста и веб-страниц для пользовательского интерфейса и кода надстройки в соответствующих папках на сервере.

Для всех типов надстроек (контентных надстроек, надстроек Outlook и надстроек области задач, а также команд надстроек) необходимо развертывать файлы веб-страниц на веб-сервере или в службе веб-хостинга, например [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).

[!include[HTTPS guidance](../includes/https-guidance.md)]

> [!TIP]
> Когда вы создаете и отлаживаете надстройку в Visual Studio, эта система развертывает и запускает соответствующие файлы веб-страниц локально с помощью IIS Express. Использовать дополнительный веб-сервер не требуется.

Для надстройок контента и области задач в поддерживаемых Office клиентских приложениях — Excel, PowerPoint, Project или Word — вам также потребуется каталог приложений в SharePoint для отправки XML-файла манифеста надстройки или развертывание надстройки с помощью [интегрированных приложений](/microsoft-365/admin/manage/test-and-deploy-microsoft-365-apps).[](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md)

Чтобы протестировать и запустить надстройку Outlook, учетная запись электронной почты Outlook пользователя должна находиться в Exchange 2013 г. или более поздней, которая доступна через Microsoft 365, Exchange Online или с помощью локальной установки. Пользователь или администратор устанавливают файлы манифестов надстроек Outlook на соответствующем сервере.

> [!NOTE]
> Учетные записи POP и IMAP в Outlook не поддерживают надстройки Office.

## <a name="client-requirements-windows-desktop-and-tablet"></a>Требования к клиенту: компьютеры и планшеты под управлением Windows

Следующее программное обеспечение необходимо для разработки надстройки Office для поддерживаемых Office или веб-клиентов, которые работают на Windows настольных, ноутбуках или планшетных устройствах.

- Для настольных компьютеров под управлением 32- и 64-разрядных версий Windows, а также таких планшетов, как Surface Pro:
  - 32- или 64-разрядная версия Office 2013 или более поздняя версия в Windows 7 или более поздней версии.
  - Excel 2013, Outlook 2013, PowerPoint 2013, Project профессиональный 2013, Project 2013 с пакетом обновления 1 (SP1), Word 2013 или более поздняя версия клиента Office, если надстройка Office тестируется или запускается специально для одного из этих клиентов Office. Клиенты Office для настольных ПК можно устанавливать локально или на клиентском компьютере с помощью технологии "нажми и работай".

  Если у вас есть Microsoft 365 подписка и у вас нет доступа к клиенту Office, вы можете скачать и установить последнюю версию [Office](https://support.microsoft.com/office/4414eaaf-0478-48be-9c42-23adc4716658).

- Браузер Internet Explorer 11 или Microsoft Edge (в зависимости от версий Windows и Office) должен быть установлен, но может не использоваться по умолчанию. Для поддержки надстроек Office клиент Office, выступающий в роли ведущего приложения, использует компоненты браузера, которые входят в состав Internet Explorer 11 или Microsoft Edge. Дополнительные сведения см. в статье [Браузеры, используемые надстройками Office](browsers-used-by-office-web-add-ins.md).

  > [!NOTE]
  > Для работы веб-надстроек Office необходимо отключить конфигурацию усиленной безопасности Internet Explorer (ESC). Если вы используете компьютер с Windows Server в качестве клиента при разработке надстроек, учитывайте, что конфигурация ESC включена по умолчанию в Windows Server.

- По умолчанию используется один из следующих браузеров: Internet Explorer 11 или последняя версия Microsoft Edge, Chrome, Firefox или Safari (Mac OS).
- Редактор HTML и JavaScript, например "Блокнот", [Visual Studio и Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs) или стороннее средство веб-разработки.

## <a name="client-requirements-os-x-desktop"></a>Требования к клиенту: настольный компьютер OS X

Outlook Mac, который распространяется в Microsoft 365, поддерживает Outlook надстройки. Запуск Outlook в Outlook Mac имеет те же требования, что и Outlook на Самом Mac: операционная система должна быть как минимум OS X v10.10 "Yosemite". Outlook для Mac использует WebKit в качестве обработчика макетов для преобразования страниц надстройки, поэтому дополнительные зависимости от браузеров отсутствуют.

Ниже приведены минимальные версии клиентов Office для Mac, которые поддерживают надстройки Office.

- Word версии 15.18 (160109)
- Excel версии 15.19 (160206)
- PowerPoint версии 15.24 (160614)

## <a name="client-requirements-browser-support-for-office-web-clients-and-sharepoint"></a>Требования к клиенту: поддержка браузеров для веб-клиентов Office в Интернете и SharePoint

Любой браузер, кроме Internet Explorer, который поддерживает ECMAScript 5.1, HTML5 и CSS3, например Microsoft Edge, Chrome, Firefox или Safari (Mac OS).

## <a name="client-requirements-non-windows-smartphone-and-tablet"></a>Требования к клиентам: Windows и планшет

Специально для Outlook на смартфонах и Windows планшетных устройствах требуется следующее программное обеспечение для тестирования и запуска Outlook надстройки.

| Приложение Office | Устройство | Операционная система | Учетная запись Exchange | Мобильный браузер |
|:-----|:-----|:-----|:-----|:-----|
|Outlook для Android|- Android-планшеты<br>- Android-смартфоны|- Android 4.4 KitKat или более поздней версии|О последнем обновлении Приложения Microsoft 365 для бизнеса или Exchange Online|Браузер не применим. Используйте родной приложение для Android. <sup>1</sup>|
|Outlook для iOS|- iPad планшеты<br>- iPhone смартфоны|- iOS 11 или более поздней|О последнем обновлении Приложения Microsoft 365 для бизнеса или Exchange Online|Браузер не применим. Используйте родной приложение для iOS. <sup>1</sup>|
|Outlook в Интернете (современный)<sup>2</sup>|- iPad 2 или более поздней<br>- Android-планшеты |- iOS 5 или более поздней<br>- Android 4.4 KitKat или более поздней версии|В Microsoft 365 Exchange Online|- Microsoft Edge<br>- Chrome<br>- Firefox<br>- Safari|
|Outlook в Интернете (классическая версия)|- iPhone 4 или более поздней<br>- iPad 2 или более поздней<br>- iPod Touch 4 или более поздней|- iOS 5 или более поздней|Локальное Exchange Server 2013 или более позднее|- Safari|

> [!NOTE]
> <sup>1</sup> OWA для Android, OWA для iPad и OWA для iPhone для родных приложений были [обесценена](https://support.microsoft.com/office/076ec122-4576-4900-bc26-937f84d25a4b).
>
> <sup>2</sup> Современные Outlook в Интернете на iPhone и Android-смартфонах больше не требуются или доступны для тестирования Outlook надстройки.

[!INCLUDE [How to distinguish between classic and modern Outlook on the web](../includes/classic-versus-modern-Outlook-on-the-web.md)]

## <a name="see-also"></a>См. также

- [Обзор платформы надстроек Office](../overview/office-add-ins.md)
- [Доступность клиентских приложений и платформ Office для надстроек Office](../overview/office-add-in-availability.md)
- [Браузеры, используемые надстройками Office](browsers-used-by-office-web-add-ins.md)
