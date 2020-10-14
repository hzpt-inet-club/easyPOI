package com.inet.codebase.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelTarget;
import com.baomidou.mybatisplus.annotation.TableName;
import com.baomidou.mybatisplus.annotation.IdType;
import com.baomidou.mybatisplus.annotation.TableId;
import java.time.LocalDateTime;
import java.io.Serializable;
import java.util.Date;

import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;

/**
 * <p>
 * 
 * </p>
 *
 * @author HCY
 * @since 2020-10-13
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
@TableName("tbl_user")
@ExcelTarget("users")
public class User implements Serializable {

    private static final long serialVersionUID = 1L;

    /**
     * 用户序号
     */

    @TableId(value = "user_id", type = IdType.AUTO)
    private Integer userId;

    /**
     * 用户姓名
     */
    @Excel(name = "姓名")
    private String userName;

    /**
     * 用户生日
     */
    @Excel(name = "生日" , importFormat = "yyyy-MM-dd HH:mm:ss")
    @JsonFormat(pattern = "yyyy-MM-dd HH:mm:ss", timezone = "GMT+8")
    private Date userBirthday;

    /**
     * 用户爱好
     */
    @Excel(name = "爱好")
    private String userHabby;

    /**
     * 用户身份证
     */
    @Excel(name = "身份证")
    private String userIdentity;

    /**
     * 用户住址
     */
    @Excel(name = "住址")
    private String userAddress;


}
