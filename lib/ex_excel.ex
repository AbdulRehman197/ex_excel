defmodule ExExcel do
  use Rustler, otp_app: :ex_excel, crate: "excelNif"

  # When your NIF is loaded, it will override this function.
  def xlsx(_path, _sheet_name), do: :erlang.nif_error(:nif_not_loaded)
  def ods(_path, _sheet_name), do: :erlang.nif_error(:nif_not_loaded)
  def xlsx_sheet_names(_path), do: :erlang.nif_error(:nif_not_loaded)
  def ods_sheet_names(_path), do: :erlang.nif_error(:nif_not_loaded)
end