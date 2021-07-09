package com.cssca.cloud;

/**
 * Created by devina(luoxiang) on 2021/1/14.
 */
public class Constants {

	public  final class TCS_RESPONSE_CODE {
		public static final String G90000000="G90000000";
		public static final String G90000013="G90000013";
		public static final String G90000002="G90000002";

	}



	public static final String E_INVOICE_SUBSCRIPTION="eInvoiceSubscription";


	/**
	 * 设备绑定通知
	 */
	public static final String TCS_CMD_DEVICE_BIND="device_bind";

	/**
	 * 纳税人核定信息变更
	 */
	public static final String TCS_CMD_TAXPAYER_APPROVE="taxpayer_approve";

	/**
	 * 票种、税种、税目、计税公式等基础信息变更
	 */
	public static final String TCS_CMD_TAX_META_DATA="tax_data";

	/**
	 * 汇率信息更新
	 */
	public static final String TCS_CMD_EXCHANGE_RATE="exchange_rate";

	/**
	 * 发票作废/负数发票审核结果
	 */
	public static final String TCS_CMD_INV_APPROVE="inv_approve";

	/**
	 * 站内信
	 */
	public static final String TCS_CMD_NOTICE="notice";

}
