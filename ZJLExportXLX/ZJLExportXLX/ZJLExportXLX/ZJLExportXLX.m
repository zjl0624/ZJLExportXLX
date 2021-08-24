//
//  ZJLExportXLX.m
//  ZJLExportXLX
//
//  Created by zjl on 2021/8/23.
//

#import "ZJLExportXLX.h"

@implementation ZJLExportXLX
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
@end
