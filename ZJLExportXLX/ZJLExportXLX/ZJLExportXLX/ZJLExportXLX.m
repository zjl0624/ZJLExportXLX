//
//  ZJLExportXLX.m
//  ZJLExportXLX
//
//  Created by zjl on 2021/8/23.
//

#import "ZJLExportXLX.h"
#import "xlsxwriter.h"

static id _instance;
@interface ZJLExportXLX ()

@end
static lxw_workbook  *workbook;
static lxw_worksheet *worksheet;
@implementation ZJLExportXLX
+ (instancetype)sharedInstance
{
    static dispatch_once_t onceToken;
    dispatch_once(&onceToken, ^{
        _instance = [[self alloc] init];
    });
    return _instance;
}

+ (void)createXLXWithTitleArray:(NSArray *)titleArray contentArray:(NSArray<NSArray *> *)contentArray {
    // 创建存放XLS文件数据的数组
    NSMutableArray  *xlsDataMuArr = [[NSMutableArray alloc] init];
    [titleArray enumerateObjectsUsingBlock:^(id  _Nonnull obj, NSUInteger idx, BOOL * _Nonnull stop) {
        [xlsDataMuArr addObject:obj];
        if (idx == titleArray.count - 1) {
            [xlsDataMuArr addObject:@"\n"];
        }else {
            [xlsDataMuArr addObject:@"\t"];
        }
    }];
    
    [contentArray enumerateObjectsUsingBlock:^(NSArray * _Nonnull obj, NSUInteger idx, BOOL * _Nonnull stop) {
        [obj enumerateObjectsUsingBlock:^(id  _Nonnull obj1, NSUInteger idx1, BOOL * _Nonnull stop1) {
            [xlsDataMuArr addObject:obj1];
            if (idx1 == obj.count - 1) {
                [xlsDataMuArr addObject:@"\n"];
            }else {
                [xlsDataMuArr addObject:@"\t"];
            }
        }];
    }];
    
    NSString *muStr = [xlsDataMuArr componentsJoinedByString:@""];
    // 文件管理器
    NSFileManager *fileManager = [[NSFileManager alloc]init];
    //使用UTF16才能显示汉字；如果显示为#######是因为格子宽度不够，拉开即可
    NSData *fileData = [muStr dataUsingEncoding:NSUTF16StringEncoding];
    // 文件路径
    NSString *path = NSHomeDirectory();
    NSString *filePath = [path stringByAppendingPathComponent:@"/Documents/export.xlsx"];
    NSLog(@"文件路径：\n%@",filePath);
    // 生成xls文件
    [fileManager createFileAtPath:filePath contents:fileData attributes:nil];
}


- (void)createXLXByLibXlsxwriterWithTitleArray:(NSArray *)titleArray contentArray:(NSArray<NSArray *> *)contentArray {
    // 文件路径
    NSString *path = NSHomeDirectory();
    NSString *filePath = [path stringByAppendingPathComponent:@"/Documents/export.xlsx"];
    NSLog(@"文件路径：\n%@",filePath);
    workbook  = workbook_new([filePath UTF8String]);// 创建新xlsx文件，路径需要转成c字符串
    worksheet = workbook_add_worksheet(workbook, NULL);// 创建sheet
    worksheet_set_column(worksheet,COLS("A:A"), 50, NULL);//设置列宽
    // 添加格式
    lxw_format *titleFormat = workbook_add_format(workbook);
    format_set_bold(titleFormat);
    format_set_font_color(titleFormat, LXW_COLOR_RED);
    format_set_align(titleFormat, LXW_ALIGN_CENTER);// 水平居中
    format_set_align(titleFormat, LXW_ALIGN_VERTICAL_CENTER);//垂直居中
    
    
    [titleArray enumerateObjectsUsingBlock:^(id  _Nonnull obj, NSUInteger idx, BOOL * _Nonnull stop) {
        worksheet_write_string(worksheet, 0, idx, [obj UTF8String], titleFormat);
    }];
    
    [contentArray enumerateObjectsUsingBlock:^(NSArray * _Nonnull obj, NSUInteger idx, BOOL * _Nonnull stop) {
        [obj enumerateObjectsUsingBlock:^(id  _Nonnull obj1, NSUInteger idx1, BOOL * _Nonnull stop1) {
            worksheet_write_string(worksheet, (int)idx + 1, idx1, [obj1 UTF8String], NULL);
        }];
    }];
    //关闭，保存文件
    workbook_close(workbook);
}
@end
