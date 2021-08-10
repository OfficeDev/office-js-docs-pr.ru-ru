---
title: Тестирование и отладка надстроек Office
description: Узнайте, как тестировать и отлаживать свою надстройку Office
ms.date: 05/19/2021
localization_priority: Priority
ms.openlocfilehash: bf3fb2dc869b382212c607144b5c61728e393c24c0ec5397dd035c6906830ac3
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57094160"
---
# <a name="test-and-debug-office-add-ins"></a>Тестирование и отладка надстроек Office

Эта статья содержит рекомендации по тестированию, отладке и диагностике надстроек Office.

## <a name="test-cross-platform-and-for-multiple-versions-of-office"></a>Тестирование кроссплатформенных выпусков и выпусков для нескольких версий Office

Надстройки Office запускаются на основных платформах, поэтому требуется протестировать надстройку на всех платформах, где ваши пользователи могут запускать Office. К ним обычно относятся Office в Интернете, Office для Windows (как подписка, так и единовременная покупка), Office для Mac, Office для iOS и (для надстроек Outlook) Office для Android. Однако могут возникать ситуации, когда вы точно знаете, что никто из ваших пользователей не будет работать на некоторых платформах. Например, если вы создаете надстройку для компании, которая требует, чтобы пользователи работали на компьютерах с Windows и подпиской на Office, вам не нужно выполнять тестирование в Office для Mac или единовременно приобретенных экземплярах для Windows.

> [!NOTE]
> На компьютерах с Windows браузер, используемый надстройкой, определяется версией Windows и Office. Дополнительные сведения см. в статье [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Надстройки, предлагаемые через AppSource, проходят проверку, включающую тестирование на всех платформах. Кроме того, надстройки тестируются в Office для Интернета со всеми основными современными браузерами, включая Microsoft Edge (WebView2 на основе Chromium), Chrome и Safari. Соответственно, перед отправкой в AppSource необходимо протестировать эти платформы и браузеры. Дополнительные сведения о проверке см. в статье [Политики сертификации коммерческой платформы Marketplace](/legal/marketplace/certification-policies), особенно в [разделе 1120.3](/legal/marketplace/certification-policies#11203-functionality), а также на странице [Доступность и применение надстроек Office](../overview/office-add-in-availability.md).
>
> AppSource не использует Internet Explorer или устаревшую версию Microsoft Edge (WebView1) для тестирования надстроек в Office для Интернета. Но если многие ваши пользователи будут применять эти два браузера для открытия Office в Интернете, вам следует протестировать их. Дополнительные сведения см. в статье [Поддержка Internet Explorer 11](../develop/support-ie-11.md) и [Устранение проблем с Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues). Office по-прежнему поддерживает эти браузеры для надстроек, поэтому если вы считаете, что столкнулись с ошибкой при работе надстройки в них, создайте проблему в репозитории [office-js](https://github.com/OfficeDev/office-js/issues/new/choose).

## <a name="sideload-an-office-add-in-for-testing"></a>Загрузка неопубликованной надстройки Office для тестирования

Вы можете установить надстройку Office для тестирования, не размещая ее в каталоге надстроек. Процедура отличается для разных платформ, а в некоторых случаях и для разных продуктов. Следующие статьи посвящены загрузке неопубликованных надстроек Office на определенной платформе или в определенном продукте.

- [Загрузка неопубликованных надстроек Office в Windows](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [Загрузка неопубликованных надстроек Office в Office в Интернете](sideload-office-add-ins-for-testing.md)

- [Загрузка неопубликованных надстроек Office на iPad и Mac](sideload-an-office-add-in-on-ipad-and-mac.md)

- [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="debug-an-office-add-in"></a>Отладка надстройки Office

Процедура отладки также отличается для разных платформ. Следующие статьи посвящены отладке надстроек Office на определенной платформе.

- [Подключение отладчика из области задач (в Windows)](attach-debugger-from-task-pane.md)

- [Отладка надстроек с помощью средств разработчика F12 в Windows 10](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

- [Отладка надстроек в Office в Интернете](debug-add-ins-in-office-online.md)

- [Отладка надстроек Office на iPad и Mac](debug-office-add-ins-on-ipad-and-mac.md)

- [Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"](debug-with-vs-extension.md)

## <a name="validate-an-office-add-in-manifest"></a>Проверка манифеста надстройки Office

Информацию о проверке манифеста надстройки Office и устранении связанных с ним неполадок см. в [этой статье](troubleshoot-manifest.md).

## <a name="troubleshoot-user-errors"></a>Устранение ошибок, с которыми сталкиваются пользователи

Информацию об устранении основных ошибок, с которыми сталкиваются пользователи при работе с надстройками Office, см. в [этой статье](testing-and-troubleshooting.md).
