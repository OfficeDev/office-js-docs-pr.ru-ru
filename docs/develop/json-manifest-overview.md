---
title: Манифест Teams для надстроек Office (предварительная версия)
description: Ознакомьтесь с предварительной версией манифеста JSON.
ms.date: 06/15/2022
ms.localizationpriority: high
ms.openlocfilehash: c739ace05992812e0de733edea2f60cf393f3c48
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659642"
---
# <a name="teams-manifest-for-office-add-ins-preview"></a>Манифест Teams для надстроек Office (предварительная версия)

Корпорация Майкрософт вносит ряд улучшений в платформу разработчиков Microsoft 365. Эти улучшения повышают согласованность при разработке, развертывании, установке и администрировании всех типов расширений Microsoft 365, включая надстройки Office. Эти изменения совместимы с существующими надстройками. 

Одним из важных улучшений, над которыми мы работаем, является возможность создания единого блока распространения для всех расширений Microsoft 365 с помощью формата манифеста и схемы на основе текущего манифеста Teams в формате JSON.

Мы выполнили важный первый шаг к достижению этих целей, сделав возможным создание надстроек Outlook, работающих только в Windows, с версией JSON-манифеста Teams.

> [!NOTE]
> Новый манифест доступен в предварительной версии и может быть изменен на основе отзывов. Мы рекомендуем опытным разработчикам надстроек поэкспериментировать с ним. Манифест предварительной версии не следует использовать в рабочих надстройках. 

В предварительной версии применяются следующие ограничения.

- Предварительная версия манифеста Teams поддерживает только надстройки Outlook и только в подписке Office для Windows. Мы работаем над расширением поддержки на Excel, PowerPoint и Word.
- Пока невозможно объединить и загрузить неопубликованную надстройку в приложении Teams, например в личной вкладке Teams или в других типах расширений Microsoft 365. В ближайшие месяцы мы продолжим расширение предварительной версии для поддержки этих сценариев и предоставим дополнительные средства для обновления манифестов до формата предварительной версии.

> [!TIP]
> Готовы приступить к работе с предварительной версией манифеста Teams? Начните со статьи [Создание надстройки Outlook с помощью манифеста Teams (предварительная версия)](../quickstarts/outlook-quickstart-json-manifest.md).

## <a name="overview-of-the-json-manifest"></a>Общие сведения о манифесте JSON

### <a name="schemas-and-general-points"></a>Схемы и общие точки

Существует только одна схема для [манифеста JSON предварительной версии](/microsoftteams/platform/resources/dev-preview/developer-preview-intro) в отличие от текущего XML-манифеста, который содержит в общей сложности семь [схем](/openspecs/office_file_formats/ms-owemxml/c6a06390-34b8-4b42-82eb-b28be12494a8).  

### <a name="conceptual-mapping-of-the-preview-json-and-current-xml-manifests"></a>Концептуальное сопоставление манифеста JSON предварительной версии и текущего XML-манифеста

В этом разделе описывается манифест JSON предварительной версии для читателей, знакомых с текущим XML-манифестом. Некоторые моменты, которые следует учитывать: 

- JSON не различает значение атрибута и элемента (в отличие от XML). Обычно JSON, сопоставляемый с XML-элементом, превращает значение элемента и каждый атрибут в дочернее свойство. В следующем примере показана разметка XML и ее эквивалент JSON.
  
  ```xml
  <MyThing color="blue">Some text</MyThing>
  ```

  ```json
  "myThing" : {
      "color": "blue",
      "text": "Some text"
  }
  ```
- В текущем XML-манифесте существует множество мест, где элемент с именем во множественном числе содержит дочерние элементы с отдельной версией того же имени. Например, разметка для настройки пользовательского меню включает элемент **\<Items\>**, который может содержать несколько дочерних элементов **\<Item\>**. Эквивалент JSON этих элементов во множественном числе — это свойство с массивом в качестве значения. Элементы массива являются *анонимными* объектами, а не свойствами с именами item или item1, item2 и т. д. Ниже приведен пример.

  ```json
  "items": [
      {
          -- markup for a menu item is here --
      },
      {
          -- markup for another menu item is here --
      }
  ]
  ```

#### <a name="top-level-structure"></a>Структура верхнего уровня

Корневой уровень манифеста JSON предварительной версии, который примерно соответствует элементу **\<OfficeApp\>** в текущем XML-манифесте, является анонимным объектом. 

Дочерние объекты элемента **\<OfficeApp\>** обычно делятся на две условные категории. Элемент **\<VersionOverrides\>** относится к одной категории. Другая состоит из всех остальных дочерних объектов элемента **\<OfficeApp\>**, которые в совокупности называются базовым манифестом. Манифест JSON предварительной версии имеет аналогичное деление. Существует свойство extension верхнего уровня, назначение и дочерние свойства которого приблизительно соответствуют элементу **\<VersionOverrides\>**. Манифест JSON предварительной версии также содержит более 10 других свойств верхнего уровня, которые совместно имеют то же назначение, что и базовый манифест XML-манифеста. Эти другие свойства можно рассматривать в совокупности как базовый манифест JSON-манифеста. 

> [!NOTE]
> Когда появится возможность объединить надстройку с другими типами расширений Microsoft 365 в одном манифесте, станут доступны другие свойства верхнего уровня, которые не соответствуют концепции базового манифеста. Обычно свойство верхнего уровня предусматривается для каждого типа расширения Microsoft 365, например configurableTabs, bots и connectors. Примеры см. в [документации по манифесту Teams](/microsoftteams/platform/resources/schema/manifest-schema). Эта структура четко указывает, что свойство extension представляет надстройку Office как один тип расширения Microsoft 365.

#### <a name="base-manifest"></a>Базовый манифест

Свойства базового манифеста указывают характеристики надстройки, которые должен содержать *любой* тип расширения Microsoft 365. Это относится ко вкладкам и расширениям для сообщений Teams, а не только к надстройкам Office. К этим характеристикам относятся общедоступное имя и уникальный идентификатор. В следующей таблице показано сопоставление некоторых критически важных свойств верхнего уровня в манифесте JSON предварительной версии с XML-элементами в текущем манифесте, где принцип сопоставления является *назначением* разметки.

|Свойство JSON|Назначение|Элементы XML|Комментарии|
|:-----|:-----|:-----|:-----|
|"$schema"| Определяет схему манифеста. | атрибуты **\<OfficeApp\>** и **\<VersionOverrides\>** | |
|"id"| GUID надстройки. | **\<Id\>**| |
|"version"| Версия надстройки. | **\<Version\>** | |
|"manifestVersion"| Версия схемы манифеста. |  атрибуты **\<OfficeApp\>** | |
|"name"| Общедоступное имя надстройки. | **\<DisplayName\>** | |
|"description"| Общедоступное описание надстройки.  | **\<Description\>** | |
|"accentColor"||| Это свойство не имеет аналога в текущем XML-манифесте и не используется в предварительной версии манифеста JSON. Но оно должно присутствовать. |
|"developer"| Определяет разработчика надстройки. | **\<ProviderName\>** | |
|"localizationInfo"| Настраивает языковой стандарт по умолчанию и другие поддерживаемые языковые стандарты. | **\<DefaultLocale\>** и **\<Override\>** | |
|"webApplicationInfo"| Определяет веб-приложение надстройки по его имени в Azure Active Directory. | **\<WebApplicationInfo\>** | В текущем XML-манифесте элемент **\<WebApplicationInfo\>** находится внутри **\<VersionOverrides\>**, а не в базовом манифесте. |
|"authorization"| Определяет все разрешения Microsoft Graph, необходимые надстройке. | **\<WebApplicationInfo\>** | В текущем XML-манифесте элемент **\<WebApplicationInfo\>** находится внутри **\<VersionOverrides\>**, а не в базовом манифесте. |

Элементы **\<Hosts\>**, **\<Requirements\>** и **\<ExtendedOverrides\>** являются частью базового манифеста в текущем XML-манифесте. Но концепции и назначение, связанные с этими элементами, настраиваются внутри свойства extension в манифесте JSON предварительной версии. 

#### <a name="extension-property"></a>Свойство extension

Свойство extension в манифесте JSON предварительной версии в основном представляет характеристики надстройки, которые не имеют отношения к другим типам расширений Microsoft 365. Например, приложения Office, которые расширяет надстройка (такие как Excel, PowerPoint, Word и Outlook), указаны в свойстве extension, как и настройки ленты приложения Office. Назначение конфигурации свойства extension очень похоже на назначение элемента **\<VersionOverrides\>** в текущем XML-манифесте.

> [!NOTE]
> Раздел **\<VersionOverrides\>** текущего XML-манифеста содержит систему "двойного перехода" для многих строковых ресурсов. Строки, включая URL-адреса, указаны с присвоением идентификатора в дочернем элементе **\<Resources\>** объекта **\<VersionOverrides\>**. Элементы, для которых требуется строка, содержат атрибут `resid`, соответствующий идентификатору строки в элементе **\<Resources\>**. Свойство extension манифеста JSON предварительной версии упрощает процесс, определяя строки непосредственно в качестве значений свойств. В манифесте JSON нет элементов, аналогичных элементу **\<Resources\>**.

В следующей таблице показано сопоставление некоторых высокоуровневых дочерних свойств свойства extension в манифесте JSON предварительной версии с XML-элементами в текущем манифесте. Точечная нотация используется для ссылки на дочерние свойства.

|Свойство JSON|Назначение|Элементы XML|Комментарии|
|:-----|:-----|:-----|:-----|
| "requirements.capabilities" | Определяет наборы обязательных элементов, которые требуются надстройке, чтобы ее можно было установить. | **\<Requirements\>** и **\<Sets\>** | |
| "requirements.scopes" | Определяет приложения Office, в которых можно установить надстройку. | **\<Hosts\>** |  |
| "ribbons" | Ленты, которые настраивает надстройка. | **\<Hosts\>**, **ExtensionPoints** и различные элементы **\*FormFactor** | Свойство ribbons представляет собой массив анонимных объектов, каждый из которых объединяет назначение этих трех элементов. См. раздел [Таблица "ribbons"](#ribbons-table).|
| "alternatives" | Указывает обратную совместимость с эквивалентной надстройкой COM, XLL или обоими вариантами. | **\<EquivalentAddins\>** | Базовые сведения см. в разделе [EquivalentAddins — дополнительные сведения](/javascript/api/manifest/equivalentaddins#see-also). |
| "runtimes"  | Настраивает различные виды надстроек без пользовательского интерфейса, например пользовательские надстройки только для функций и функции, запускаемые непосредственно из пользовательских кнопок ленты. | **\<Runtimes\>**. **\<FunctionFile\>** и **\<ExtensionPoint\>** (типа CustomFunctions) |  |
| "autoRunEvents" | Настраивает обработчик для указанного события. | **\<Event\>** и **\<ExtensionPoint\>** (типа Events) |  |

##### <a name="ribbons-table"></a>Таблица "ribbons"

В следующей таблице дочерние свойства анонимных дочерних объектов в массиве "ribbons" сопоставлены с XML-элементами текущего манифеста. 

|Свойство JSON|Назначение|Элементы XML|Комментарии|
|:-----|:-----|:-----|:-----|
| "contexts" | Указывает поверхности команд, которые настраивает надстройка. | различные элементы **\*CommandSurface**, например **PrimaryCommandSurface** и **MessageReadCommandSurface** |  |
| "tabs" | Настраивает пользовательские вкладки ленты. | **\<CustomTab\>** | Имена и иерархия свойств-потомков объекта tabs соответствуют потомкам объекта **\<CustomTab\>**.  |

## <a name="sample-preview-json-manifest"></a>Пример манифеста JSON предварительной версии

Ниже приведен пример JSON-манифеста предварительной версии для надстройки.

```json
{
  "$schema": "https://raw.githubusercontent.com/OfficeDev/microsoft-teams-app-schema/op/extensions/MicrosoftTeams.schema.json",
  "id": "00000000-0000-0000-0000-000000000000",
  "version": "1.0.0",
  "manifestVersion": "devPreview",
  "name": {
    "short": "Name of your app (<=30 chars)",
    "full": "Full name of app, if longer than 30 characters (<=100 chars)"
  },
  "description": {
    "short": "Short description of your app (<= 80 chars)",
    "full": "Full description of your app (<= 4000 chars)"
  },
  "icons": {
    "outline": "outline.png",
    "color": "color.png"
  },
  "accentColor": "#230201",
  "developer": {
    "name": "Contoso",
    "websiteUrl": "https://www.contoso.com",
    "privacyUrl": "https://www.contoso.com/privacy",
    "termsOfUseUrl": "https://www.contoso.com/servicesagreement"
  },
  "localizationInfo": {
    "defaultLanguageTag": "en-us",
    "additionalLanguages": [
      {
        "languageTag": "es-es",
        "file": "es-es.json"
      }
    ]
  },
  "webApplicationInfo": {
    "id": "00000000-0000-0000-0000-000000000000",
    "resource": "api://www.contoso.com/prodapp"
  },
  "authorization": {
    "permissions": {
      "resourceSpecific": [
        {
          "name": "Mailbox.ReadWrite.User",
          "type": "Delegated"
        }
      ]
    }
  },
  "extensions": [
    {
      "requirements": {
        "scopes": [ "mail" ],
        "capabilities": [
          {
            "name": "Mailbox", "minVersion": "1.1"
          }
        ]
      },
      "runtimes": [
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.10"
              }
            ]
          },
          "id": "eventsRuntime",
          "type": "general",
          "code": {
            "page": "https://contoso.com/events.html",
            "script": "https://contoso.com/events.js"
          },
          "lifetime": "short",
          "actions": [
            {
              "id": "onMessageSending",
              "type": "executeFunction"
            },
            {
              "id": "onNewMessageComposeCreated",
              "type": "executeFunction"
            }
          ]
        },
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.1"
              }
            ]
          },
          "id": "commandsRuntime",
          "type": "general",
          "code": {
            "page": "https://contoso.com/commands.html",
            "script": "https://contoso.com/commands.js"
          },
          "lifetime": "short",
          "actions": [
            {
              "id": "action1",
              "type": "executeFunction"
            },
            {
              "id": "action2",
              "type": "executeFunction"
            },
            {
              "id": "action3",
              "type": "executeFunction"
            }
          ]
        }
      ],
      "ribbons": [
        {
          "contexts": [
            "mailCompose"
          ],
          "tabs": [
            {
              "builtInTabId": "TabDefault",
              "groups": [
                {
                  "id": "dashboard",
                  "label": "Controls",
                  "controls": [
                    {
                      "id": "control1",
                      "type": "button",
                      "label": "Action 1",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "Action 1 Title",
                        "description": "Action 1 Description"
                      },
                      "actionId": "action1"
                    },
                    {
                      "id": "menu1",
                      "type": "menu",
                      "label": "My Menu",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "My Menu",
                        "description": "Menu with 2 actions"
                      },
                      "items": [
                        {
                          "id": "menuItem1",
                          "type": "menuItem",
                          "label": "Action 2",
                          "supertip": {
                            "title": "Action 2 Title",
                            "description": "Action 2 Description"
                          },
                          "actionId": "action2"
                        },
                        {
                          "id": "menuItem2",
                          "type": "menuItem",
                          "label": "Action 3",
                          "icons": [
                            {
                              "size": 16,
                              "file": "test_16.png"
                            },
                            {
                              "size": 32,
                              "file": "test_32.png"
                            },
                            {
                              "size": 80,
                              "file": "test_80.png"
                            }
                          ],
                          "supertip": {
                            "title": "Action 3 Title",
                            "description": "Action 3 Description"
                          },
                          "actionId": "action3"
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "contexts": [ "mailRead" ],
          "tabs": [
            {
              "builtInTabId": "TabDefault",
              "groups": [
                {
                  "id": "dashboard",
                  "label": "Controls",
                  "controls": [
                    {
                      "id": "control1",
                      "type": "button",
                      "label": "Action 1",
                      "icons": [
                        {
                          "size": 16,
                          "file": "test_16.png"
                        },
                        {
                          "size": 32,
                          "file": "test_32.png"
                        },
                        {
                          "size": 80,
                          "file": "test_80.png"
                        }
                      ],
                      "supertip": {
                        "title": "Action 1 Title",
                        "description": "Action 1 Description"
                      },
                      "actionId": "action1"
                    }
                  ]
                }
              ]
            }
          ]
        }
      ],
      "autoRunEvents": [
        {
          "requirements": {
            "capabilities": [
              {
                "name": "MailBox", "minVersion": "1.10"
              }
            ]
          },
          "events": [
            {
              "type": "newMessageComposeCreated",
              "actionId": "onNewMessageComposeCreated"
            },
            {
              "type": "messageSending",
              "actionId": "onMessageSending",
              "options": {
                "sendMode": "promptUser"
              }
            }
          ]
        }
      ],
      "alternates": [
        {
          "requirements": {
            "scopes": [ "mail" ]
          },
          "prefer": {
            "comAddin": {
              "progId": "ContosoExtension"
            }
          },
          "hide": {
            "storeOfficeAddin": {
              "officeAddinId": "00000000-0000-0000-0000-000000000000",
              "assetId": "WA000000000"
            }
          }
        }
      ]
    }
  ]
}
```

## <a name="next-steps"></a>Дальнейшие действия

- [Создание надстройки Outlook с помощью манифеста Teams (предварительная версия)](../quickstarts/outlook-quickstart-json-manifest.md)