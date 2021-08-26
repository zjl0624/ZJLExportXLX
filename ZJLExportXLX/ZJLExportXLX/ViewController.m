//
//  ViewController.m
//  ZJLExportXLX
//
//  Created by zjl on 2021/8/23.
//

#import "ViewController.h"
#import "ZJLExportXLX.h"
#import "xlsxwriter.h"

@interface ViewController ()

@end

@implementation ViewController

- (void)viewDidLoad {
    [super viewDidLoad];
    self.view.backgroundColor = [UIColor whiteColor];
    NSArray *titleArray = @[@"时间",@"地点",@"Person"];
    NSArray *contentArray = @[@[@"2016-12-06 17:18:40",@"广州",@"张三"],@[@"2016-12-07 17:18:40",@"成都",@"李四"],@[@"2016-12-08 17:18:40",@"广州",@"王麻子"]];
//    [ZJLExportXLX createXLXWithTitleArray:titleArray contentArray:contentArray];
    [[ZJLExportXLX sharedInstance] createXLXByLibXlsxwriterWithTitleArray:titleArray contentArray:contentArray];
    
}

@end
