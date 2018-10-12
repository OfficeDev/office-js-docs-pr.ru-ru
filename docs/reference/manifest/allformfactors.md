# <a name="allformfactors-element"></a>Элемент AllFormFactors

Указывает параметры всех форм-факторов для надстройки. В настоящее время только в настраиваемых функциях применяется **AllFormFactors**. Элемент** AllFormFactors** является обязательным при использовании настраиваемых функций.

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  [ExtensionPoint](extensionpoint.md) |  Да |  Определяет, где предоставляются функции надстройки. |

## <a name="allformfactors-example"></a>Пример использования AllFormFactors

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
