//
//  ZJLExportXLX.h
//  ZJLExportXLX
//
//  Created by zjl on 2021/8/23.
//

#import <Foundation/Foundation.h>

NS_ASSUME_NONNULL_BEGIN

@interface ZJLExportXLX : NSObject
+ (instancetype)sharedInstance;
+ (void)createXLXWithTitleArray:(NSArray *)titleArray contentArray:(NSArray<NSArray *> *)contentArray;
- (void)createXLXByLibXlsxwriterWithTitleArray:(NSArray *)titleArray contentArray:(NSArray<NSArray *> *)contentArray;

@end

NS_ASSUME_NONNULL_END
