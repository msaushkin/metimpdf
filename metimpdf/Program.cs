using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.xml.xmp;
using iTextSharp.xmp;
using iTextSharp.xmp.impl;
using iTextSharp.xmp.options;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace metimdf{
    class Program{
        // Перечисление для удобного формирования типов записываемого значения
        enum ValueType{
            Bag,
            Seq,
            Alt,
            Property
        }
        // вспомогательный класс для создания карты 
        class MapObject{
            // конструктор класса без параметров
            public MapObject(){ }
            // конструкто класса без параметра сокращенного имени(для инициализации в случае если другое имя такое же как и оригинальное)
            public MapObject(string schemaNs, string name, ValueType valueType, bool isSingle = true) : this(schemaNs, name, name, valueType, isSingle){ }
            /**
             * конструктор со всеми парамерами 
             * schemaNs - схема принадлежности свойства (prism, dc, xmp, xmpRights и т.д. все они (кроме prism) есть в объекте констант XmpConst)
             * name - полное имя свойства (совпадает с именем в data файле)
             * valueName - сконвертированное имя (на случай если имя в выходном файле отличается от оригинального имени)
             * valueType - тип свойства (Объект перечисление ValueType, один из трёх типов массива или просто однострочное свойство)
             * isSingle - работает только если объект массив элементов и отвечает за то, следует ли распилить элемент и вставить его коллекцией или одной строкой
             */
            public MapObject(string schemaNs, string name, string valueName, ValueType valueType, bool isSingle = true){
                XmlSchemaNs = schemaNs;
                Name = name;
                ValueName = valueName;
                ValueType = valueType;
                IsSingle = isSingle;
            }

            public string XmlSchemaNs { get; set; }
            public string Name { get; set; }
            public string ValueName { get; set; }
            public ValueType ValueType { get; set; }

            public bool IsSingle { get; set; }
        }
        
        static void Main(string[] args){
        	if (args.Length != 3){
				System.Console.WriteLine("The correct number of arguments is 3");
				System.Console.WriteLine("Description:");
				System.Console.WriteLine("Metimpdf is the program for [met]adata [im]port into a [pdf] file");
				System.Console.WriteLine("as a part of the Math-Net.Ru Meta Data Creator Tools.");
				System.Console.WriteLine("Use:");
				System.Console.WriteLine("metimpdf <pdf file> <metadata> <output pdf file>");
				System.Console.WriteLine("Example:");
				System.Console.WriteLine("metimpdf source.pdf meta.data output.pdf");
				Console.Write("Press any key to continue . . . ");
				Console.ReadKey(true);
				return;
        	}
        	
			if (args.Length == 3){
        		//System.Console.WriteLine(args[0]);
				//System.Console.WriteLine(args[1]);
				//System.Console.WriteLine(args[2]);
				try{
					string SourcePdfFile = args[0];
					string MetadataFile  = args[1];
					string OutputPdfFile = args[2];
				
					string prismSchema = "http://prismstandard.org/namespaces/basic/2.2/";
            	
					// Создаем и регистрируем несуществующую в системе схему prism
            		XmpMetaFactory.SchemaRegistry.RegisterNamespace(prismSchema, "prism");

            		/** 
             		* Создаём карту для корректной конвертации и вставки значений в metadata
             		* т.к. ключевые слова указанные в .data файле не существуют в пространствах
             		*/
            		List<MapObject> map = new List<MapObject>();
            		map.Add(new MapObject(XmpConst.NS_DC, "Title", "title", ValueType.Alt));
            		map.Add(new MapObject(XmpConst.NS_DC, "Author", "creator", ValueType.Seq, false));
            		map.Add(new MapObject(XmpConst.NS_DC, "Keywords", "subject", ValueType.Bag, false));
            		map.Add(new MapObject(XmpConst.NS_DC, "Doi", "identifier", ValueType.Property));
            		map.Add(new MapObject(XmpConst.NS_DC, "Publisher", "publisher", ValueType.Bag));
            		map.Add(new MapObject(XmpConst.NS_DC, "Rights", "rights", ValueType.Alt));
            		map.Add(new MapObject(XmpConst.NS_DC, "Description", "description", ValueType.Alt));
            	
            		map.Add(new MapObject(XmpConst.NS_PDF, "Keywords", ValueType.Property));
            		map.Add(new MapObject(XmpConst.NS_PDF, "Producer", ValueType.Property));
            	
					map.Add(new MapObject(XmpConst.NS_XMP, "CreatorTool", ValueType.Property));
	            
					map.Add(new MapObject(XmpConst.NS_XMP_RIGHTS, "Rights", "UsageTerms", ValueType.Alt));
					map.Add(new MapObject(XmpConst.NS_XMP_RIGHTS, "Marked", "Marked", ValueType.Property));
            		map.Add(new MapObject(XmpConst.NS_XMP_RIGHTS, "RightsUrl", "WebStatement", ValueType.Property));

            		map.Add(new MapObject(prismSchema, "Doi", "doi", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "DoiUrl", "url", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "Issn", "issn", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "eIssn", "eissn", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "Volume", "volume", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "Number", "number", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "CoverDisplayDate", "coverDisplayDate", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "CoverDate", "coverDate", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "IssueName", "issueName", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "PageRange", "pageRange", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "StartingPage", "startingPage", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "EndingPage", "endingPage", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "AggregationType", "aggregationType", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "Platform", "originPlatform", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "PublicationName", "publicationName", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "Edition", "edition", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "Section", "section", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "Subsection1", "subsection1", ValueType.Property));
            		map.Add(new MapObject(prismSchema, "Rights", "copyright", ValueType.Property));
            		

            		// Открываем поток для чтения/записи значений
            		using (FileStream fs = new FileStream("metadata.xmp", FileMode.Create, FileAccess.ReadWrite)){
                		// Считываем все стрроки из .data файла
                		string[] lines = File.ReadAllLines(MetadataFile);

                		// Создаем словарь разбора строк в сопоставлении ключ:значение
                		Dictionary<string, string> info = new Dictionary<string, string>();
                	
                		// цикл перебора строк для разбора в словарь ключ:значение
                		foreach(string line in lines){
                    		string[] splitted = line.Split('='); // делим строку попалам относительно знака =
                    		string key = splitted[0].Trim(); // ключ(значение слева от = )
                    		string value = splitted[1].Trim(); // значение(значение справа от =)
                    		value = value.Substring(1, value.Length - 2); // убираем скобки {}
                    		info.Add(key, value);
                		}

                		/**
                 		* Создаём объект для корректной современной записи Xmp и в качестве значения конструктора 
                 		* передаём только поток
                 		*/
                		XmpWriter xmpWriter = new XmpWriter(fs);
                
                		
                		xmpWriter.XmpMeta.SetProperty(prismSchema, "platform", "print");
                		xmpWriter.XmpMeta.SetProperty(prismSchema, "edition", "My print");
                		xmpWriter.XmpMeta.SetProperty(prismSchema, "section", "My print");
                		xmpWriter.XmpMeta.SetProperty(prismSchema, "subsection1", "My print");
                		
                		
                		// цикл перебора словаря для дальнейшего разбора
                		foreach (string key in info.Keys){
                			// цикл перебора карты на поиск соответствующих ключевым словам объектов
                    		foreach(MapObject p in map){
                				if(p.Name == key){ // если имя объекта равно ключевому слову из data то
									/** 
                             		* сохраняем значение в xmpWriter с помощь специального 
                             		* упращенного метода сохранения(сам метод создан в конце класса)
                             		*/
                             		SetValue(xmpWriter, p, info[key]);
                				}
                			}
                		}
                		

                		
                		
                		// модификатор языка
                		XmpNode languageQualifier = new XmpNode("xml:lang", "x-default", new PropertyOptions {Qualifier = true });
                		
                		// более низкий уровень взаимодействия с деревом xml для проставки дополнительных атрибутов
                		XmpNode groupNode = (xmpWriter.XmpMeta as XmpMetaImpl).Root.FindChildByName(XmpConst.NS_DC);
                
                		// находим элементы которым нужно указать языковую конструкцию
                		XmpNode titleDcNode = groupNode.FindChildByName("dc:title");
                		XmpNode rightsDcNode = groupNode.FindChildByName("dc:rights");
                		XmpNode descriptyopmDcNode = groupNode.FindChildByName("dc:description");

                		// проставляем модификатор
                		titleDcNode.GetChild(1).AddQualifier(languageQualifier);
                		rightsDcNode.GetChild(1).AddQualifier(languageQualifier);
                		descriptyopmDcNode.GetChild(1).AddQualifier(languageQualifier);

                		groupNode = (xmpWriter.XmpMeta as XmpMetaImpl).Root.FindChildByName(XmpConst.NS_XMP_RIGHTS);
                		XmpNode rightsXmpRightsNode = groupNode.FindChildByName("xmpRights:UsageTerms");
                		rightsXmpRightsNode.GetChild(1).AddQualifier(languageQualifier);
                		
                		// Закрываем xmpWriter
                		xmpWriter.Close();
            		} // после этой скобочки файл metadata.xmp сохранится и запишется автоматически

            		// Внедрение metadata.xmp в SourcePdfFile и вывод в OutputPdfFile 
            		PdfReader reader = new PdfReader(SourcePdfFile);
            		PdfStamper stamper = new PdfStamper(reader, new FileStream(OutputPdfFile, FileMode.Create));
            		using(FileStream fs = File.OpenRead("metadata.xmp")){
            			byte[] buffer = new byte[fs.Length];
                		fs.Read(buffer, 0, buffer.Length);
                		stamper.XmpMetadata = buffer;
            		}

            		stamper.Close();
            		reader.Close();			
			}
			// Catching iTextSharp.text.DocumentException if any
			catch (DocumentException de){
				throw de;
				}
			// Catching System.IO.IOException if any
			catch (IOException ioe){
				throw ioe;
				}
			}			

			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
        }

        /**
         * Метод упрощающий запись в xmp файл, т.к. из Java перекочевал ключевой недостаток библиотеки: 4 разных метода для записи свойства объектов
         * в качестве аргементов принимает
         * 1 Объект XmpWriter созданный выше
         * 2 Объект карты из map
         * 3 значение для записи
         */
        static void SetValue(XmpWriter writer, MapObject mapObject, string value){
        	// Проверяет что значение НЕ записывается одной строкой и что это не требуемое значения не типа "Свойство"
            if (!mapObject.IsSingle && mapObject.ValueType != ValueType.Property){
                // делит значения по точке с запятой или запятой
                string[] values = value.Split(new string[] { " ; ", "," }, StringSplitOptions.RemoveEmptyEntries);
                // перебирает только что разделенную строку и последовательно записывает в коллекцию соответствующим способом
                foreach(string v in values){
                    PropertyOptions arrayOptions = new PropertyOptions();
                    
                    if (mapObject.ValueType == ValueType.Bag)
                        arrayOptions.Array = true;
                    if(mapObject.ValueType == ValueType.Alt)
                        arrayOptions.ArrayAlternate = true;
                    if (mapObject.ValueType == ValueType.Seq)
                        arrayOptions.ArrayOrdered = true;

                    writer.XmpMeta.AppendArrayItem(mapObject.XmlSchemaNs, mapObject.ValueName, arrayOptions, v.Trim(), new PropertyOptions { HasLanguage = true, HasQualifiers = true });
                }
            }
            else{
                if (mapObject.ValueType != ValueType.Property){
                    PropertyOptions arrayOptions = new PropertyOptions();
                    if (mapObject.ValueType == ValueType.Bag)
                        arrayOptions.Array = true;
                    if (mapObject.ValueType == ValueType.Alt)
                        arrayOptions.ArrayAlternate = true;
                    if (mapObject.ValueType == ValueType.Seq)
                        arrayOptions.ArrayOrdered = true;

                    writer.XmpMeta.AppendArrayItem(mapObject.XmlSchemaNs, mapObject.ValueName, arrayOptions, value.Trim(), new PropertyOptions { HasLanguage=true });
                }
                else{
                    writer.XmpMeta.SetProperty(mapObject.XmlSchemaNs, mapObject.ValueName, value);
                }
            }
        }
    }

}
