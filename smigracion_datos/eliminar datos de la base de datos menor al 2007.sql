-- ****************************************** Elimina todos los datos *****************************************
--		                        use sigh
--					DELETE FROM FARMINVENTARIOCABECERA
--					DELETE FROM FARMINVENTARIODETALLE
--					DELETE FROM FARMINVENTARIO
--					DELETE FROM FARMMOVIMIENTONOTAINGRESO
--					DELETE FROM FARMMOVIMIENTOPROGRAMAS
--					DELETE FROM FacturacionBienesPagos
--					DELETE FROM FactOrdenesBienes
--					DELETE FROM FacturacionBienesFinanciamientos
--					DELETE FROM FARMMOVIMIENTOVENTASDETALLE
--					DELETE FROM FARMMOVIMIENTOVENTAS
--					DELETE FROM FARMMOVIMIENTODETALLE
--					DELETE FROM FARMMOVIMIENTO
--					DELETE FROM FARMPREVENTADETALLE
--					DELETE FROM FARMPREVENTA
--					DELETE FROM FARMSALDODETALLADO
--					DELETE FROM FARMSALDO
--					DELETE FROM Proveedores
-- 		delete from ImagMovimientoImagenes
-- 		delete from ImagMovimientoIngresos
-- 		delete from ImagMovimientoSalidas
-- 		delete from ImagMovimientoDetalle
-- 		delete from ImagMovimiento
 --					DELETE FROM FacturacionServicioDevoluciones
--					DELETE FROM FacturacionServicioFinanciamientos
--					DELETE FROM FacturacionServicioDespacho
--					DELETE FROM FacturacionServicioPagos
--					DELETE FROM FactOrdenServicioPagos
--					DELETE FROM FactOrdenServicio
--                                      delete from CajaComprobantesPago
--		use sigh
--              Delete from AtencionesNacimientos 
--              Delete from AtencionesEmergencia 
  --            Delete from AtencionesEstanciaHospitalaria 
--              Delete from AtencionesDiagnosticos 
--              delete from AtencionesConvenio
--              delete from AtencionesInterconsultas
--              delete from    FacturacionCuentasAtencion
--              DELETE FROM citas
--              DELETE FROM HistoriasSolicitadas
--              DELETE FROM MovimientosHistoriaClinica
--              Delete from ATENCIONES 
--             delete from citasbloqueadas
--             delete from ProgramacionMedica
--             delete from CajaGestion
-- use sigh
-- Delete from HistoriasClinicas
--Delete from camasMovimientos
--Delete from camas
--Delete from auditoria
--delete from cajagestion
--delete from pacientes
--delete from MedicosEspecialidad
--delete from medicos
--delete from EstablecimientosNoMinsa
--delete from usuariosRoles where idEmpleado<>738
--delete from ArchiveroServicio where idEmpleado<>738
--delete from empleados where idEmpleado<>738
--delete from reporte
-- ****************************************** Elimina todos los datos *****************************************



-- ****************************************** Elimina datos menores al 2007*****************************************
--Antes tiene que relacionar "marcar DELETE"   historiasSolicitadas VS movimientosHistoriaClinica
DELETE FROM HistoriasSolicitadas WHERE (YEAR(FechaSolicitud) < 2007)
DELETE FROM MovimientosHistoriaClinica WHERE     (YEAR(FechaMovimiento) < 2007)

--Antes tiene que relacionar "actualizar en cascada los campos relacionados"   factOrdenesBienesInsumo VS facturacionBienesInsumo
--Antes tiene que relacionar "actualizar en cascada los campos relacionados"   factOrdenesBienesInsumo VS facturacionBienesInsumo
DELETE FROM factOrdenesBienesInsumo
WHERE     (YEAR(fechaCreacion) < 2007)

--Antes tiene que relacionar "actualizar en cascada los campos relacionados"   citas VS programacionmedica
--Antes tiene que relacionar "actualizar en cascada los campos relacionados"   citas VS programacionmedica
DELETE FROM citas
WHERE     (YEAR(fecha) < 2007)
--Antes tiene que relacionar "actualizar en cascada los campos relacionados"   facturacioncuentasatencion VS cajacomprobantepago
--Antes tiene que relacionar "actualizar en cascada los campos relacionados"   facturacioncuentasatencion VS cajacomprobantepago
DELETE FROM facturacioncuentasatencion
WHERE     (YEAR(fechaapertura) < 2007)


--Antes tiene que relacionar "actualizar en cascada los campos relacionados"   factordenesservicion VS facturacionservicio
--Antes tiene que relacionar "actualizar en cascada los campos relacionados"   factordenesservicio VS facturacionservicio
DELETE FROM factordenesservicio
WHERE     (YEAR(fechaorden) < 2007)






--Procedimiento que elimina ATENCIONES-........


        declare @lnIdAtencion int
        declare tmpAtenciones cursor for select idAtencion from ATENCIONES where YEAR(FechaIngresO)<2007
        open tmpAtenciones
        fetch next from tmpAtenciones into @lnIdAtencion
        while (@@fetch_status<>-1)
        begin
              Delete from AtencionesNacimientos where idAtencion=@lnIdAtencion
              Delete from AtencionesEmergencia where idAtencion=@lnIdAtencion
              Delete from AtencionesEstanciaHospitalaria where idAtencion=@lnIdAtencion
              Delete from AtencionesDiagnosticos where idAtencion=@lnIdAtencion             
              Delete from FacturacionServicios where idAtencion=@lnIdAtencion
              Delete from FactOrdenesServicio where idAtencion=@lnIdAtencion   
              Delete from atenciones where idAtencion=@lnIdAtencion   
              fetch next from tmpAtenciones into @lnIdAtencion
        end




--Procedimiento que elimina PACIENTES ..............443475
   declare @id int
        declare tmpPacientes cursor for select idPaciente from PACIENTES 
        open tmpPacientes
        fetch next from tmpPacientes into @id
        while (@@fetch_status<>-1)
        begin 
              declare tmpAtenciones cursor for select idPaciente from Atenciones  where idPaciente=@id
              open tmpAtenciones
              fetch next from tmpAtenciones
              if @@fetch_status=-1
              begin
                   PRINT @id
                   Delete from Pacientes where idPaciente=@id
                   Delete from HistoriasClinicas where idPaciente=@id
              end  
              close tmpAtenciones
              DEALLOCATE tmpAtenciones
              fetch next from tmpPacientes into @id
        end
-- ****************************************** Elimina datos menores al 2007*****************************************
