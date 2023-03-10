EasyPOI (Thư viện xử lý Excel và Word cho ứng dụng Java)
===========================
EasyPOI là thư viện tích hợp ApachePOI với mục tiêu hướng đến giúp giảm thiểu nỗ lực xử lý xuất/nhập nội dung Exel &
Word. Bộ thư viện cung cấp hệ thống Anotation giúp bạn dễ dàng sử dụng tính năng thông qua vài dòng code đơn giản.

Thư viện được xây dựng dựa trên AutoPOI

Phiên bản： 1.0.0（09/03/2022）
 
---------------------------
Điểm nổi trội
--------------------------

	1.Đơn giản dễ sử dụng
	2.Cung cấp nhiều giao diện, dễ dàng bổ sung thêm khi cần thiết
	3.Tích hợp một cách nhanh chóng
	4.Hỗ trợ AbstractView cho web

---------------------------
Một số class hay dùng
---------------------------

	1.ExcelExportUtil xuất dữ liệu ra định dạng Excel (dạng Basic hoặc sử dụng Template)
	2.ExcelImportUtil nhập dữ liệu từ định dạng Exel
	3.WordExportUtil  xuất dữ liệu ra định dạng Word (chỉ định dạng docx)

	
---------------------------
Sự khác biệt giữa xuất XLS và XLSX Exel
---------------------------

	1. Thời gian xuất XLS nhanh hơn 2-3 lần so với XLSX
	2. Kích thước xuất XLS gấp 2-3 lần hoặc nhiều hơn so với XLSX
	3. Tốc độ xuất dựa trên chất lượng mạng internet và khả năng xử lý của serve

	
---------------------------
Mô tả cấu trúc Source code
---------------------------

	1.vts-kit-lib-easy-poi: project cha để dễ quản lý 
	2.easypoi: module core xử lý việc xuất/nhập Exel & Word
	3.viewpoi: module spring-mvc dựa trên AbstractView, đơn giản hóa việc xuất/nhập

--------------------------
Cách tích hợp
--------------------------

```xml

<dependency>
    <groupId>com.atviettelsolutions</groupId>
    <artifactId>viewpoi</artifactId>
    <version>1.0.0</version>
</dependency>
```

--------------------------
Các hàm so sánh được hỗ trợ
--------------------------

- Spacesplitting
- Trinocular operation {{test ? obj:obj2}}
- n: indicates This cell is a numeric type {{n:}}
- le: stands for length {{le:()}} in ifelse using {{le:() > 8 ? obj1 : obj2}}
- fd: format time {{fd:(obj; yyyy-MM-dd)}}
- fn: format number {{fn:(obj;.00)}}
- fe: Traverse data, create row
- !fe: Traverse data without creating row
- fe: Move down insert, move the current row, the following row down, all the .size() rows, and insert
- !if: Delete the current column {{!if:(test)}}
- single quotes indicate constant values "1", then the output is 1

---------------------------
Code Example
---------------------------
1.Sử dụng Anotation

```Java

@ExcelTarget("courseEntity")
public class CourseEntity implements java.io.Serializable {
    /** primaryKey */
    private String id;
    /** courseName */
    @Excel(name = "Họ và tên", orderNum = "1", needMerge = true)
    private String name;
    /** mathTeacher primary key */
    @ExcelEntity(id = "mathTeacher")
    @ExcelVerify()
    private TeacherEntity mathTeacher;
    /** physicalTeacher primary key */
    @ExcelEntity(id = "physicalTeacher")
    private TeacherEntity physicalTeacher;

    @ExcelCollection(name = "Danh sách học sinh", orderNum = "4")
    private List<StudentEntity> students;
}
```

2. Xuất dữ liệu cơ bản với tham số đầu vào

```
    HSSFWorkbook workbook=ExcelExportUtil.exportExcel(new ExportParams(
        "2412312","test","test"),CourseEntity.class,list);
```
3. 

```
    ExportParams params = new ExportParams("2412312", "tes", "test");
	params.setAddIndex(true);
	HSSFWorkbook workbook = ExcelExportUtil.exportExcel(params,
			TeacherEntity.class, telist);
```
4. Xuất dữ liệu cơ bản với tham số đầu vào

```
  List<ExcelExportEntity> entity = new ArrayList<ExcelExportEntity>();
	entity.add(new ExcelExportEntity("Tên", "name"));
	entity.add(new ExcelExportEntity("Giới tính", "sex"));

	List<Map<String, String>> list = new ArrayList<Map<String, String>>();
	Map<String, String> map;
	for (int i = 0; i < 10; i++) {
		map = new HashMap<String, String>();
		map.put("name", "1" + i);
		map.put("sex", "2" + i);
		list.add(map);
	}

	HSSFWorkbook workbook = ExcelExportUtil.exportExcel(new ExportParams(
			"test", "test"), entity, list);	
```
5. Xuất Theo cấu hình template

```
    TemplateExportParams params = new TemplateExportParams();
    params.setHeadingRows(2);
    params.setHeadingStartRow(2);
    Map<String,Object> map = new HashMap<String, Object>();
    map.put("year", "2023");
    map.put("sunCourses", list.size());
    Map<String,Object> obj = new HashMap<String, Object>();
    map.put("obj", obj);
    obj.put("name", list.size());
    params.setTemplateUrl("com/viettel/vtskit/easypoi/template/exportTemp.xls");
    Workbook book = ExcelExportUtil.exportExcel(params, CourseEntity.class, list,map);
```
6. Xuất với file template chỉ định

```
    ImportParams params = new ImportParams();
	params.setTitleRows(2);
	params.setHeadRows(2);
	//params.setSheetNum(9);
	params.setNeedSave(true);
	long start = new Date().getTime();
	List<CourseEntity> list = ExcelImportUtil.importExcel(new File(
			"d:/tt.xls"), CourseEntity.class, params);
```

7. Xuất dữ liệu sử dụng SpringMVC view
	
```
	@RequestMapping(value = "/exportXls")
	public ModelAndView exportXls(HttpServletRequest request, HttpServletResponse response) {
		ModelAndView mv = new ModelAndView(new JeecgEntityExcelView());
		List<EasyDemo> pageList = easyDemoService.list();
		mv.addObject(NormalExcelConstants.FILE_NAME,"my_export_data");
		mv.addObject(NormalExcelConstants.CLASS,Student.class);
		mv.addObject(NormalExcelConstants.PARAMS,new ExportParams("my_title", "my_sheet_name"));
		mv.addObject(NormalExcelConstants.DATA_LIST,pageList);
		return mv;
	}
```

| Các loại view          | Chức năng   | Mô tả                          |
|------------------------|-------------|--------------------------------|
| JeecgMapExcelView      | Dạng xem xuất đối tượng thực thể    | Ví dụ: List<EasyDemo>         |
| JeecgEntityExcelView   | Dạng xem xuất đối tượng Map   | List<Map<String, String>> list |
| JeecgTemplateExcelView | Chế độ xem xuất mẫu Excel | -                              | 
| JeecgTemplateWordView  | Chế độ xem xuất mẫu Word  | -                              |

8. Validate dữ liệu đầu vào

```
/**
 * Email Object
 */
@Excel(name = "Email", width = 25)
@ExcelVerify(isEmail = true, notNull = true)
private String email;
/**
 * phone number 
 */
@Excel(name = "Mobile", width = 20)
@ExcelVerify(isMobile = true, notNull = true)
private String mobile;

ExcelImportResult<ExcelVerifyEntity> result=ExcelImportUtil.importExcelVerify(new File("d:/tt.xls"),ExcelVerifyEntity.class,params);
        for(int i=0;i<result.getList().size();i++){
        System.out.println(ReflectionToStringBuilder.toString(result.getList().get(i)));
        }
```

9.Nhập dữ liêệu

```
    ImportParams params=new ImportParams();
        List<Map<String, Object>>list=ExcelImportUtil.importExcel(new File(
        "d:/tt.xls"),Map.class,params);
```

11.Cách sử dụng bảng từ điển. 
Với: 
- dictTable: là tên bảng cơ sở dữ liệu, 
- dicCode: là tên trường liên kết 
- dicText: là trường tương ứng với nội dung được hiển thị trong file excel

```
    @Excel(name = "department", dictTable = "t_s_depart", dicCode = "id", dicText = "departname")
    private java.lang.String depart;
```

12. Sử dụng Replace để thay thế dữ liệu.
Ví dụ: Nếu dữ liệu đầu vào là 0/1 ，thì trong file Exel sẽ là Nam/Nữ

```
    @Excel(name = "testSubstitution", width = 15, replace = {"man_1", "woman_0"})
    private java.lang.String fdReplace;
```

13. Sử dụng Convert để bổ sung dữ liệu/thay thế dữ liệu

- exportConvert：thay thế giá trị dữ liệu lúc xuất dữ liệu.
- importConvert：thay thế giá trị dữ liệu lúc nhập dữ liệu.

```
    @Excel(name = "Test the conversion", width = 15, exportConvert = true, importConvert = true)
    private java.lang.String fdConvert;
    /**
    * Example of conversion values： The field value is concated to "meta"
    * @return
    */
    public String convertgetFdConvert(){
        return this.fdConvert+"meta";
        }

    /**
    * Example of converting values: Replace "meta" in excel cells
    * @return
    */
    public void convertsetFdConvert(String fdConvert){
        this.fdConvert=fdConvert.replace("meta","");
    }
```

---------------------------
Chú thích Anotation
---------------------------

@Excel

| Thuộc tính     | kiểu     | giá trịn mặc định | Chức năng                                                                                                                                                                              |
|----------------|----------|-------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| name           | String   | null              | Column name, which supports name id                                                                                                                                                    |
| needMerge      | boolean  | fasle             | Merge cells vertically (for multiple rows created by merging lists with single cells in a list)                                                                                        |
| orderNum       | String   | "0"               | Sorting of columns, support for name id                                                                                                                                                |
| replace        | String[] | {}                | It is worth replacing the export is {a_id, b_id} import in reverse                                                                                                                     |
| savePath       | String   | "upload"          | Import file save path, if it is a picture can be filled in, the default is uploadclassName IconEntity, this class corresponds to uploadIcon                                            |
| type           | int      | 1                 | Export type 1 is text 2 is picture, 3 is function, 10 is number. The default is text                                                                                                   |
| width          | double   | 10                | Column width                                                                                                                                                                           |
| height         | double   | 10                | Column height, later plan to use the @Excel Target's height, this will be abandoned, attention                                                                                         |
| isStatistics   | boolean  | fasle             | Automatic statistics, in appending a row of statistics, putting all data and output This processing will engulf the exception, please note this                                        |
| isHyperlink    | boolean  | FALSE             | If a hyperlink is required, an interface return object needs to be implemented                                                                                                         |
| isImportField  | boolean  | TRUE              | Check the field, see if this field is imported into Excel, if there is no description is wrong Excel, read failed, support name id                                                     |
| exportFormat   | String   | ""                | The exported time format, whether this is empty to determine whether the date needs to be formatted                                                                                    |
| importFormat   | String   | ""                | The imported time format, whether this is empty or not, determines whether the date needs to be formatted                                                                              |
| format         | String   | ""                | The time format is equivalent to setting both exportFormat and importFormat                                                                                                            |
| databaseFormat | String   | "yyyyMMddHHmmss"  | Export time settings, if the field is of type Date, you do not need to set the database If it is a string type, this database format needs to be set to convert the time format output |
| numFormat      | String   | ""                | Number formatting, the parameter is Pattern, and the object used is Decimal Format                                                                                                     |
| imageType      | int      | 1                 | Export type 1 Read from file 2 Read from database Default is file Similarly, import is the same                                                                                        |
| suffix         | String   | ""                | Text suffix, e.g. %90 becomes 90%                                                                                                                                                      |
| isWrap         | boolean  | TRUE              | wrapping is supported \n                                                                                                                                                               |
| mergeRely      | int[]    | {}                | Merge cell dependencies, e.g. the second column merge is based on the first column, then {0} will do                                                                                                                                                       |
| mergeVertical  | boolean  | fasle             | Merge cells with the same content vertically                                                                                                                                                                           |
| fixedIndex     | int      | -1                | Corresponding to the Excel column, ignore the name                                                                                                                                                                         |
| isColumnHidden | boolean  | FALSE             | Export hidden columns                                                                                                                                                                                  |

@ExcelCollection

| attribute       | type       | default value             | function               |
|----------|----------|-----------------|------------------|
| id       | String   | null            | Define the ID             |
| name     | String   | null            | Define collection column names and support nanm ids |
| orderNum | int      | 0               | Sorting, support name id     |
| type     | Class<?> | ArrayList.class | Use when importing objects that are created        |

Single table export entity annotation source code

```
public class SysUser implements Serializable {

    /**id*/
    private String id;

    /**username */
    @Excel(name = "Tên đăng nhập", width = 15)
    private String username;

    /**realname*/
    @Excel(name = "Họ và tên", width = 15)
    private String realname;

    /**avatar*/
    @Excel(name = "Ảnh đại diện", width = 15)
    private String avatar;

    /**birthday*/
    @Excel(name = "Ngày sinh", width = 15, format = "yyyy-MM-dd")
    private Date birthday;

    /****sex**（1：man 2：woman）*/
    @Excel(name = "Giới tính", width = 15, dicCode = "sex")
    private Integer sex;

    /**email*/
    @Excel(name = "Email", width = 15)
    private String email;

    /**phone*/
    @Excel(name = "Điện thoại", width = 15)
    private String phone;

    /**status(1：Hoạt động  0：Không hoạt động）*/
    @Excel(name = "Trạng thái", width = 15, replace = {"Hoạt động_1", "Không hoạt động_0"})
    private Integer status;
```

One-to-many export entity annotation source code

```Java

@Data
public class EasyDemoOrderMainPage {

    /**primaryKey*/
    private java.lang.String id;
    /**Order number*/
    @Excel(name = "Order number", width = 15)
    private java.lang.String orderCode;
    /**Order type*/
    private java.lang.String ctype;
    /**Order date*/
    @Excel(name = "Order date", width = 15, format = "yyyy-MM-dd")
    private java.util.Date orderDate;
    /**Order amount*/
    @Excel(name = "Order amount", width = 15)
    private java.lang.Double orderMoney;
    /**Order notes*/
    private java.lang.String content;
    /**Created by*/
    private java.lang.String createBy;
    /**Creation time*/
    private java.util.Date createTime;
    /**Modify the person*/
    private java.lang.String updateBy;
    /**Modification time*/
    private java.util.Date updateTime;

    @ExcelCollection(name = "Client")
    private List<EasyDemoOrderCustomer> easyDemoOrderCustomerList;
    @ExcelCollection(name = "Ticket")
    private List<EasyDemoOrderTicket> easyDemoOrderTicketList;
}
```