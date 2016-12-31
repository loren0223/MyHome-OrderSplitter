package utilities;
import java.math.BigDecimal;

/**

*
�ѩ�Java��²�������������T����B�I�ƶi��B��A�o�Ӥu�������Ѻ�

* �T���B�I�ƹB��A�]�A�[����M�|�٤��J�C

*/

public class Arith
{

	//�q�{���k�B����
	private static final int DEF_DIV_SCALE = 10;

	//�o���������Ҥ�
	private Arith(){}



	/**
	* ���Ѻ�T���[�k�B��C
	* @param v1 �Q�[��
	*
	@param v2 �[��
	* @return ��ӰѼƪ��M
	*/
	
	public static double add(double v1,double v2)
	{
		BigDecimal b1 = new
		BigDecimal(Double.toString(v1));
		
		BigDecimal b2 = new
		BigDecimal(Double.toString(v2));
		
		return b1.add(b2).doubleValue();
	}
	
	/**
	* ���Ѻ�T����k�B��C
	* @param v1 �Q���
	*
	@param v2 ���
	* @return ��ӰѼƪ��t
	*/
	
	public static double sub(double v1,double v2)
	{
		BigDecimal b1 = new
		BigDecimal(Double.toString(v1));
		
		BigDecimal b2 = new
		BigDecimal(Double.toString(v2));
		
		return b1.subtract(b2).doubleValue();
	}
	
	/**
	* ���Ѻ�T�����k�B��C
	* @param v1 �Q����
	*
	@param v2 ����
	* @return ��ӰѼƪ��n
	*/
	
	public static double mul(double v1,double v2) 
	{
		BigDecimal b1 = new
		BigDecimal(Double.toString(v1));
		
		BigDecimal b2 = new
		BigDecimal(Double.toString(v2));
		
		return b1.multiply(b2).doubleValue();
	}
	
	
	
	/**
	* ���ѡ]�۹�^��T�����k�B��A��o�Ͱ����ɪ����p�ɡA��T��
	*
	�p���I�H�Z10��A�H�Z���Ʀr�|�٤��J�C
	* @param v1 �Q����
	* @param v2 ����
	*
	@return ��ӰѼƪ���
	*/
	
	public static double div(double v1,double v2)
	{
		return div(v1,v2,DEF_DIV_SCALE);
	}
	
	
	
	/**
	*
	���ѡ]�۹�^��T�����k�B��C��o�Ͱ����ɪ����p�ɡA��scale�Ѽƫ�
	* �w��סA�H�Z���Ʀr�|�٤��J�C
	* @param v1
	�Q����
	* @param v2 ����
	* @param scale ��ܪ�ܻݭn��T��p���I�H�Z�X��C
	*
	@return ��ӰѼƪ���
	*/
	
	public static double div(double v1,double v2,int scale)
	{
		if(scale<0)
		{
			throw new IllegalArgumentException("The scale must be a positive integer or	zero");
		}
		
		BigDecimal b1 = new BigDecimal(Double.toString(v1));
		BigDecimal b2 = new BigDecimal(Double.toString(v2));
		
		return b1.divide(b2,scale,BigDecimal.ROUND_HALF_UP).doubleValue();
	}
	
	/**
	*
	��Ƭ۰����l��
	* @param v1 �Q����
	* @param v2 ����
	* @return �l�� 
	*/
	
	public static double mod(double v1,double v2)
	{
		BigDecimal b1 = new BigDecimal(Double.toString(v1));
		BigDecimal b2 = new BigDecimal(Double.toString(v2));
		
		BigDecimal[] x = b1.divideAndRemainder(b2);
		return x[1].doubleValue();
	}
	
	
	
	/**
	* ���Ѻ�T���p�Ʀ�|�٤��J�B�z�C
	* @param v �ݭn�|�٤��J���Ʀr
	* @param scale �p���I�Z�O�d�X��
	* @return �|�٤��J�Z�����G
	*/
	
	public static double round(double v,int scale)
	{
	
		if(scale<0)
		{
			throw new IllegalArgumentException("The scale must be a positive integer or zero");
		}
		
		BigDecimal b = new BigDecimal(Double.toString(v));
		BigDecimal one = new BigDecimal("1");
		
		return b.divide(one,scale,BigDecimal.ROUND_HALF_UP).doubleValue();
	
	}
	
	/**
	* ���Ѻ�T���p�Ʀ�L����˥h�B�z�C
	* @param v �ݭn�L����˥h���Ʀr
	* @param scale �p���I�Z�O�d�X��
	* @return �L����˥h�����G
	*/
	
	public static double rounddown(double v,int scale)
	{
	
		if(scale<0)
		{
			throw new IllegalArgumentException("The scale must be a positive integer or zero");
		}
		
		BigDecimal b = new BigDecimal(Double.toString(v));
		BigDecimal one = new BigDecimal("1");
		
		return b.divide(one,scale,BigDecimal.ROUND_DOWN).doubleValue();
	
	}
	
	/**
	* ���Ѻ�T���p�Ʀ�L����i��B�z�C
	* @param v �ݭn�L����i�쪺�Ʀr
	* @param scale �p���I�Z�O�d�X��
	* @return �L����i�쪺���G
	*/
	
	public static double roundup(double v,int scale)
	{
	
		if(scale<0)
		{
			throw new IllegalArgumentException("The scale must be a positive integer or zero");
		}
		
		BigDecimal b = new BigDecimal(Double.toString(v));
		BigDecimal one = new BigDecimal("1");
		
		return b.divide(one,scale,BigDecimal.ROUND_UP).doubleValue();
	
	}

}