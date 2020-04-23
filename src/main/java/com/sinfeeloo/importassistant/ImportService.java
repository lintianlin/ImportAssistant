package com.sinfeeloo.importassistant;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Objects;
import java.entity;

public class ImportService {
    /**
     * 批量发货
     *
     * @param file
     * @return
     * @throws Exception
     */
    public int importExcel(Integer companyId, MultipartFile file) throws Exception {
        int result = 0;
        if (!file.isEmpty()) {
            String suffix = Objects.requireNonNull(file.getOriginalFilename()).substring(file.getOriginalFilename().lastIndexOf(""));
            if (!".xlsx".equals(suffix)) {
                throw new Exception("导入的文件格式不正确");
            }
            //增加EXCEL表格处理开始
            Response<ExcelInfo> excelInfoResponse = ExcelTools.transferStreamToExcel(file.getInputStream());
            ExcelInfo excelInfo = excelInfoResponse.getResult();
            if (excelInfo == null) {
                throw new Exception(excelInfoResponse.getMsg());
            }
            List<ExcelTitleInfo> titleInfoList = excelInfo.getTitleInfo();
            if (titleInfoList == null || titleInfoList.size() == 0 || titleInfoList.size() != 3) {
                throw new Exception("导入的数据格式不正确");
            }
            List<GrouponOrder> list = new ArrayList<>();
            List<List<ExcelBodyInfo>> bodyInfoList = excelInfo.getBodyInfo();
            List<WechatAppletMessage> wechatAppletMessageList = new ArrayList<>();
            for (int i = 0; i < bodyInfoList.size(); i++) {
                List<ExcelBodyInfo> excelBodyInfoList = bodyInfoList.get(i);
                GrouponOrder grouponOrder = new GrouponOrder();
                for (int j = 0; j < 3; j++) {
                    switch (j) {
                        case 0://订单号
                            if (StringUtils.isNotEmpty(excelBodyInfoList.get(j).getValue())) {
                                grouponOrder.setOrderCode(excelBodyInfoList.get(j).getValue().trim());
                            } else {
                                grouponOrder.setOrderCode("");
                            }
                            break;
                        case 1://快递单号
                            grouponOrder.setExpressCode(excelBodyInfoList.get(j).getValue());
                            break;
                        case 2://快递公司
                            grouponOrder.setExpressCompany(excelBodyInfoList.get(j).getValue());
                            break;
                    }
                }
                grouponOrder.setOrderStatus(3);
                Date sendDate = new Date();
                grouponOrder.setSendTime(sendDate);
                list.add(grouponOrder);
            }
            //后台校验文件格式正确性
            if (list == null || list.size() == 0) {
                throw new Exception("导入的文件没有数据！");
            }
            for (int i = 0; i < list.size(); i++) {
                GrouponOrder grouponOrderInfo = list.get(i);
                String orderCode = grouponOrderInfo.getOrderCode();
                if (!StringUtils.isEmpty(orderCode)) {
                    if (StringUtils.isEmpty(grouponOrderInfo.getOrderCode())) {
                        throw new BusinessException("第" + (i + 2) + "行订单号为空");
                    }
                    if (StringUtils.isEmpty(grouponOrderInfo.getExpressCode())) {
                        throw new BusinessException("第" + (i + 2) + "行物流单号为空");
                    }
                    if (StringUtils.isEmpty(grouponOrderInfo.getExpressCompany())) {
                        throw new BusinessException("第" + (i + 2) + "行物流公司为空");
                    }
                    List<GrouponOrder> grouponOrderListByOrderCode = grouponOrderDao.selectByOrderCode(orderCode);
                    if (grouponOrderListByOrderCode == null || grouponOrderListByOrderCode.size() == 0) {
                        throw new BusinessException("第" + (i + 2) + "行" + orderCode + "订单号不存在");
                    }

//                    if (grouponOrderListByOrderCode.get(0).getOrderStatus() != 2 && grouponOrderListByOrderCode.get(0).getOrderStatus() != 3) {
//                        throw new BusinessException("第" + (i + 2) + "行订单状态错误");
//                    }
                    GrouponOrderDetailQuery orderDetailQuery = new GrouponOrderDetailQuery();
                    orderDetailQuery.createCriteria().andOrderCodeEqualTo(orderCode);
                    List<GrouponOrderDetail> orderDetailList = grouponOrderDetailDao.selectByExample(orderDetailQuery);
                    if (orderDetailList == null || orderDetailList.size() <= 0) {
                        throw new BusinessException("第" + (i + 2) + "行订单商品信息错误");
                    }

                    GrouponOrder tempGroupOrder = grouponOrderListByOrderCode.get(0);
                    grouponOrderInfo.setAppid(tempGroupOrder.getAppid());
                    grouponOrderInfo.setOpenid(tempGroupOrder.getOpenid());
                    grouponOrderInfo.setOrderStatusFromDataBase(tempGroupOrder.getOrderStatus());
                    grouponOrderInfo.setSendTimeFromDataBase(tempGroupOrder.getSendTime());
                    grouponOrderInfo.setOrderDetailList(orderDetailList);

                }

            }
            //批量更新数据
            System.out.println("总数量：" + list.size());
            for (int i = 0; i < list.size(); i++) {
                GrouponOrder grouponOrderItem = list.get(i);
                if (StringUtils.isNotBlank(grouponOrderItem.getOrderCode())) {
                    //只有待发货和已发货的才能更新
                    if (grouponOrderItem.getOrderStatusFromDataBase() == 2
                            || grouponOrderItem.getOrderStatusFromDataBase() == 21
                            || grouponOrderItem.getOrderStatusFromDataBase() == 22
                            || grouponOrderItem.getOrderStatusFromDataBase() == 23
                            || grouponOrderItem.getOrderStatusFromDataBase() == 3) {
                        int resultCount = grouponOrderDao.updateForDeliver(grouponOrderItem.getOrderStatus(), grouponOrderItem.getSendTime(), grouponOrderItem.getOrderCode(), grouponOrderItem.getExpressCode(), grouponOrderItem.getExpressCompany());
                        // 更新成功
                        //状态更新成功
                        if (resultCount > 0 && grouponOrderItem.getSendTimeFromDataBase() == null) {
                            //设置消息体
                            WechatAppletMessage wechatAppletMessage = setWechatAppletMessage(grouponOrderItem.getAppid(), grouponOrderItem.getOpenid(), grouponOrderItem.getOrderCode(), grouponOrderItem.getExpressCode(), grouponOrderItem.getExpressCompany(), grouponOrderItem.getSendTime(), grouponOrderItem.getOrderDetailList());
                            wechatAppletMessageList.add(wechatAppletMessage);
                        }
                    }
                }
            }
////            grouponOrderDao.batchUpdateForDeliver(list);
            //增加EXCEL表格处理结束
            result = 1;
            //异步发送微信订阅消息
            //解决非异步方法中调用异步方法失效的问题
            ((GrouponOrderService) AopContext.currentProxy()).sendWechatMessageAsync(companyId, wechatAppletMessageList);
            System.out.println("微信订阅消息总数量：" + wechatAppletMessageList.size());

        }
        return result;
    }
}
