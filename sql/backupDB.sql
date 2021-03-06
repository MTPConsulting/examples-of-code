USE [ITSGestion]
GO
/****** Object:  StoredProcedure [dbo].[_BackupBD]    Script Date: 05/09/2018 8:54:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		Oscar
-- Create date: 2018-Marzo
-- Description:	Backup de la base de datos de Azure
-- =============================================
ALTER PROCEDURE [dbo].[_BackupBD] 
AS
BEGIN
	-- Inicialización
	SET NOCOUNT ON
	
	DECLARE @db_id INT
	DECLARE @db_nombre NVARCHAR(128)
	DECLARE @ErrorMessage NVARCHAR(4000)

	SET @db_id = DB_ID()
	SET @db_nombre = DB_NAME()
	SET @ErrorMessage  = '?'

	--Verifica que se haya realizado la descarga de datos desde Azure
	IF (NOT EXISTS (SELECT * 
                 FROM INFORMATION_SCHEMA.TABLES 
                 WHERE TABLE_SCHEMA = 'dbo' 
                 AND  TABLE_NAME = '_ka_proveedores')) begin
		RAISERROR ('No existen las tablas con los datos descargados de Azure.', 16, 1, @db_id, @db_nombre)
		return
	end 

	--Inicia una transacción
	BEGIN TRANSACTION

	BEGIN TRY 
		-- Elimina los datos de transacciones
		print ''
		print 'Eliminando movimientos anteriores...'
		exec wp_blanqueo_bd @confirmar = 1
		print 'ok'

		SET NOCOUNT OFF
	
		--Agrega información de maestros faltante
		print ''
		print 'Actualizando maestros...'

		insert into km_tipos_iva
			select * 
			from _km_tipos_iva tmp
			where not exists(select * from km_tipos_iva i5 where i5.tp_iva=tmp.tp_iva)
	
		insert into km_tipos_pago
			select * 
			from _km_tipos_pago tmp
			where not exists(select * from km_tipos_pago i5 where i5.tp_medio_pago=tmp.tp_medio_pago)

		insert into ta_cuentas_bancarias
			select * 
			from _ta_cuentas_bancarias tmp
			where not exists(select * from ta_cuentas_bancarias i5 where i5.cd_cuenta_bancaria=tmp.cd_cuenta_bancaria)

		insert into vm_tipos_cobro
			select * 
			from _vm_tipos_cobro tmp
			where not exists(select * from vm_tipos_cobro i5 where i5.tp_medio_cobro=tmp.tp_medio_cobro)

		insert into vm_tipos_iva
			select * 
			from _vm_tipos_iva tmp
			where not exists(select * from vm_tipos_iva i5 where i5.tp_iva=tmp.tp_iva)

		insert into wm_diccionario
			select * 
			from _wm_diccionario tmp
			where not exists(select * from wm_diccionario i5 where i5.tp_elemento=tmp.tp_elemento)

		insert into wm_operadores
			select * 
			from _wm_operadores tmp
			where not exists(select * from wm_operadores i5 where i5.nm_email=tmp.nm_email)

		insert into wm_tipos_comprobante
			select * 
			from _wm_tipos_comprobante tmp
			where not exists(select * from wm_tipos_comprobante i5 where i5.cd_comprobante=tmp.cd_comprobante)

		print 'ok'

		--Insertando operaciones de tesorería
		print ''
		print 'Procesando movimientos de tesorería...'

		SET IDENTITY_INSERT dbo.ta_mov_bancarios ON
		insert into ta_mov_bancarios([nu_movbco], [cd_cuenta_bancaria], [fe_movimiento], [de_movimiento], [im_ingreso], [im_egreso])
			select [nu_movbco], [cd_cuenta_bancaria], [fe_movimiento], [de_movimiento], [im_ingreso], [im_egreso] from _ta_mov_bancarios
		SET IDENTITY_INSERT dbo.ta_mov_bancarios OFF

		SET IDENTITY_INSERT ta_imputaciones ON
		insert into ta_imputaciones ([cd_imputacion], [nm_imputacion], [tp_imputacion])
			select [cd_imputacion], [nm_imputacion], [tp_imputacion] from _ta_imputaciones
		SET IDENTITY_INSERT ta_imputaciones OFF

		print 'ok'

		--Insertando operaciones de compras
		print ''
		print 'Procesando movimientos de compras...'

		SET IDENTITY_INSERT ka_proveedores ON
		insert into ka_proveedores([cd_proveedor], [nm_proveedor], [nu_cuit], [tp_iva], [cd_imputacion])
			select [cd_proveedor], [nm_proveedor], [nu_cuit], [tp_iva], [cd_imputacion] from _ka_proveedores
		SET IDENTITY_INSERT ka_proveedores OFF
		
		SET IDENTITY_INSERT kc_comprobantes ON
		insert into kc_comprobantes([nu_interno_cbte], [tp_comprobante], [tp_factura], [nu_ptovta], [nu_comprobante], [fe_comprobante], [cd_proveedor], [de_comprobante], [de_observaciones], [im_neto], [im_neto_no_gravado], [po_iva1], [im_iva1], [po_iva2], [im_iva2], [im_percepcion_iva], [im_percepcion_iibb], [im_total], [nu_op_anticipo])
			select [nu_interno_cbte], [tp_comprobante], [tp_factura], [nu_ptovta], [nu_comprobante], [fe_comprobante], [cd_proveedor], [de_comprobante], [de_observaciones], [im_neto], [im_neto_no_gravado], [po_iva1], [im_iva1], [po_iva2], [im_iva2], [im_percepcion_iva], [im_percepcion_iibb], [im_total], [nu_op_anticipo] from _kc_comprobantes
		SET IDENTITY_INSERT kc_comprobantes OFF
		
		SET IDENTITY_INSERT kl_comprobantes ON
		insert into kl_comprobantes([nu_item_cbte], [nu_interno_cbte], [cd_imputacion], [de_imputacion], [im_total], [fe_imputacion]) 
			select [nu_item_cbte], [nu_interno_cbte], [cd_imputacion], [de_imputacion], [im_total], [fe_imputacion] from _kl_comprobantes
		SET IDENTITY_INSERT kl_comprobantes OFF
		
		SET IDENTITY_INSERT kc_pagos ON
		insert into kc_pagos([nu_interno_op], [fe_orden_pago], [cd_proveedor], [nu_comprobante_cancela], [de_orden_pago], [im_cancelado]) 
			select [nu_interno_op], [fe_orden_pago], [cd_proveedor], [nu_comprobante_cancela], [de_orden_pago], [im_cancelado] from _kc_pagos
		SET IDENTITY_INSERT kc_pagos OFF
		
		SET IDENTITY_INSERT kl_pagos ON
		insert into kl_pagos([nu_item_op], [nu_interno_op], [tp_medio_pago], [de_pago], [im_total], [cd_cuenta_bancaria], [fe_cheque]) 
			select [nu_item_op], [nu_interno_op], [tp_medio_pago], [de_pago], [im_total], [cd_cuenta_bancaria], [fe_cheque] from _kl_pagos
		SET IDENTITY_INSERT kl_pagos OFF

		print 'ok'

		--Insertando operaciones de ventas
		print ''
		print 'Procesando movimientos de ventas...'

		SET IDENTITY_INSERT va_clientes ON
		insert into va_clientes([cd_cliente], [nm_cliente], [nu_cuit], [tp_iva], [cd_imputacion])
			select [cd_cliente], [nm_cliente], [nu_cuit], [tp_iva], [cd_imputacion] from _va_clientes
		SET IDENTITY_INSERT va_clientes OFF
		
		SET IDENTITY_INSERT vc_comprobantes ON
		insert into vc_comprobantes([nu_interno_cbte], [tp_comprobante], [tp_factura], [nu_ptovta], [nu_comprobante], [fe_comprobante], [cd_cliente], [fe_vencimiento], [cd_imputacion], [de_comprobante], [de_observaciones], [im_neto], [im_iva], [im_total], [fe_imputacion], [nu_recibo_anticipo])
			select [nu_interno_cbte], [tp_comprobante], [tp_factura], [nu_ptovta], [nu_comprobante], [fe_comprobante], [cd_cliente], [fe_vencimiento], [cd_imputacion], [de_comprobante], [de_observaciones], [im_neto], [im_iva], [im_total], [fe_imputacion], [nu_recibo_anticipo] from _vc_comprobantes
		SET IDENTITY_INSERT vc_comprobantes OFF

		SET IDENTITY_INSERT vc_cobranzas ON
		insert into vc_cobranzas([nu_interno_rc], [fe_recibo], [cd_cliente], [nu_comprobante_cancela], [de_recibo], [im_cancelado])
			select [nu_interno_rc], [fe_recibo], [cd_cliente], [nu_comprobante_cancela], [de_recibo], [im_cancelado] from _vc_cobranzas
		SET IDENTITY_INSERT vc_cobranzas OFF
		
		SET IDENTITY_INSERT vl_cobranzas ON
		insert into vl_cobranzas([nu_item_rc], [nu_interno_rc], [tp_medio_cobro], [de_cobro], [im_total], [cd_cuenta_bancaria])
			select [nu_item_rc], [nu_interno_rc], [tp_medio_cobro], [de_cobro], [im_total], [cd_cuenta_bancaria] from _vl_cobranzas
		SET IDENTITY_INSERT vl_cobranzas OFF

		print 'ok'

		SET NOCOUNT OFF

		--Elimina las tablas importadas de Azure
		drop table [dbo].[_ka_proveedores]
		drop table [dbo].[_kc_comprobantes]
		drop table [dbo].[_kc_pagos]
		drop table [dbo].[_kl_comprobantes]
		drop table [dbo].[_kl_pagos]
		drop table [dbo].[_km_tipos_iva]
		drop table [dbo].[_km_tipos_pago]
		drop table [dbo].[_ta_cuentas_bancarias]
		drop table [dbo].[_ta_imputaciones]
		drop table [dbo].[_ta_mov_bancarios]
		drop table [dbo].[_va_clientes]
		drop table [dbo].[_vc_cobranzas]
		drop table [dbo].[_vc_comprobantes]
		drop table [dbo].[_vl_cobranzas]
		drop table [dbo].[_vm_tipos_cobro]
		drop table [dbo].[_vm_tipos_iva]
		drop table [dbo].[_wm_diccionario]
		drop table [dbo].[_wm_operadores]
		drop table [dbo].[_wm_tipos_comprobante]


		--Confirma la transacción
		COMMIT TRAN
		print ''
		print 'COMMITED'

	END TRY  
	BEGIN CATCH  
		--Revierte la transacción
		ROLLBACK TRAN
		print ''
		print 'ROLLBACK'
		SET @ErrorMessage  = ERROR_MESSAGE()
		RAISERROR (@ErrorMessage, 16, 1, @db_id, @db_nombre)
	END CATCH  

	--Hace backup de la BD en disco
	--DBCC SHRINKDATABASE ([ITSGestion], 5);  

	BACKUP DATABASE [ITSGestion]
	TO DISK='D:\SQL\MSSQL13.MSSQLSERVER\MSSQL\Backup\ITSGestion.bak'   
	WITH   
		DESCRIPTION = 'Backup BD sistema gestión ITSouth',
		INIT,
		SKIP; 

END
