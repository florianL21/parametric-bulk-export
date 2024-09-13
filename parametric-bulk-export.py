from .lib import fusion360utils as futil
import adsk.core
import adsk.fusion
import traceback
import csv
from pathlib import Path

COMMAND_NAME = "Parametric Export"
COMMAND_DESCRIPTION = "Bulk export meshes, changing selected parameters."
COMMAND_RESOURCES = ""
COMMAND_ID = "parametric-bulk-export"
TARGET_WORKSPACE = "FusionSolidEnvironment"
TARGET_PANEL = "SolidScriptsAddinsPanel"

CSV_EXPORT_FLAG = "Activate Export"
CSV_EXPORT_NAME = "Export Name"
CSV_SPECIAL_HEADERS = [CSV_EXPORT_NAME, CSV_EXPORT_FLAG]
_handlers: "list[adsk.core.EventHandler]" = []


class CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()
        self.app = adsk.core.Application.get()
        self.ui = self.app.userInterface
        self.design: adsk.fusion.Design = self.app.activeProduct  # type: ignore
        self.export_variations_count = 10
        self.total_columns = 2 + self.export_variations_count

    def notify(self, eventArgs: adsk.core.CommandCreatedEventArgs):
        try:
            self._notify(eventArgs)
        except Exception:
            if self.ui:
                self.ui.messageBox(
                    f"Panel command created failed:\n{traceback.format_exc()}"
                )

    def _notify(self, args: adsk.core.CommandCreatedEventArgs):
        cmd = args.command
        on_execute = CommandExecuteHandler()
        cmd.execute.add(on_execute)
        _handlers.append(on_execute)
        inputs = cmd.commandInputs

        export_file_types_group = inputs.addGroupCommandInput(
            "exportFileTypes", "File Types"
        )
        export_file_types_group.children.addBoolValueInput(
            "exportStlMeshBool", "STL", True
        )
        export_file_types_group.children.addBoolValueInput(
            "exportStepMeshBool", "Step", True
        )
        export_file_types_group.children.addBoolValueInput(
            "exportObjMeshBool", "Obj", True
        )
        export_file_types_group.children.addBoolValueInput(
            "export3mfMeshBool", "3MF", True
        )

        export_input_options_group = inputs.addGroupCommandInput(
            "exportImportFile", "Export/import"
        )

        radioButtonGroup = (
            export_input_options_group.children.addRadioButtonGroupCommandInput(
                "radioImportExport", " Import or Export "
            )
        )
        radioButtonGroup.isFullWidth = True
        radioButtonItems = radioButtonGroup.listItems
        radioButtonItems.add("Load CSV", True)
        radioButtonItems.add("Save starting point CSV", False)


class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()
        self.app = adsk.core.Application.get()
        self.ui = self.app.userInterface

    def notify(self, eventArgs: adsk.core.CommandEventArgs):
        try:
            self._notify(eventArgs)
        except Exception:
            if self.ui:
                self.ui.messageBox(
                    f"command executed failed:\n{traceback.format_exc()}"
                )

    def _notify(self, args: adsk.core.CommandEventArgs):
        inputs = args.command.commandInputs

        do_stl = bool(inputs.itemById("exportStlMeshBool").value)  # type: ignore
        do_step = bool(inputs.itemById("exportStepMeshBool").value)  # type: ignore
        do_obj = bool(inputs.itemById("exportObjMeshBool").value)  # type: ignore
        do_3mf = bool(inputs.itemById("export3mfMeshBool").value)  # type: ignore
        radioButtonGroup: adsk.core.RadioButtonGroupCommandInput = inputs.itemById(
            "radioImportExport"
        )  # type: ignore
        is_import = radioButtonGroup.selectedItem.name == "Load CSV"
        do_import_export(is_import, do_stl, do_step, do_obj, do_3mf)


def get_add_in_command_definition(ui: adsk.core.UserInterface):
    command_definitions = ui.commandDefinitions
    command_definition = command_definitions.itemById(COMMAND_ID)
    if not command_definition:
        command_definition = command_definitions.addButtonDefinition(
            COMMAND_ID, COMMAND_NAME, COMMAND_DESCRIPTION, COMMAND_RESOURCES
        )
    return command_definition


def command_control_by_id_for_panel(command_id: str):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not command_id:
        ui.messageBox("commandControl id is not specified")
        return None
    workspaces = ui.workspaces
    modeling_workspace = workspaces.itemById(TARGET_WORKSPACE)
    toolbar_panels = modeling_workspace.toolbarPanels
    toolbar_panel = toolbar_panels.itemById(TARGET_PANEL)
    toolbar_controls = toolbar_panel.controls
    toolbar_control = toolbar_controls.itemById(command_id)
    return toolbar_control


def command_definition_by_id(command_id: str):
    app = adsk.core.Application.get()
    ui = app.userInterface
    if not command_id:
        ui.messageBox("command_definition id is not specified")
        return None
    command_definitions = ui.commandDefinitions
    command_definition = command_definitions.itemById(command_id)
    return command_definition


def destroy_object(
    ui_obj: adsk.core.UserInterface,
    to_be_delete_obj: adsk.core.ToolbarControl | adsk.core.CommandDefinition,
):
    if not ui_obj and not to_be_delete_obj:
        return

    if to_be_delete_obj.isValid:
        to_be_delete_obj.deleteMe()
    else:
        ui_obj.messageBox("toBeDeleteObj is not a valid object")


def run(_):
    ui = None

    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        start_add_in(app, ui)
    except Exception:
        if ui:
            ui.messageBox("AddIn Start Failed:\n{}".format(traceback.format_exc()))
        futil.handle_error("run")


def start_add_in(app: adsk.core.Application, ui: adsk.core.UserInterface):
    command_definition = get_add_in_command_definition(ui)
    on_command_created = CommandCreatedHandler()
    command_definition.commandCreated.add(on_command_created)
    _handlers.append(on_command_created)

    workspaces = ui.workspaces
    modeling_workspace = workspaces.itemById(TARGET_WORKSPACE)
    toolbar_panels = modeling_workspace.toolbarPanels
    toolbar_panel = toolbar_panels.itemById(TARGET_PANEL)
    toolbar_controls_panel = toolbar_panel.controls
    toolbar_control_panel = toolbar_controls_panel.itemById(COMMAND_ID)
    if not toolbar_control_panel:
        toolbar_control_panel = toolbar_controls_panel.addCommand(
            command_definition, ""
        )
        toolbar_control_panel.isVisible = True
        futil.log(f"{COMMAND_ID} successfully added to add ins panel")


def do_import_export(
    isImport: bool, do_stl: bool, do_step: bool, do_obj: bool, do_3mf: bool
):
    app = adsk.core.Application.get()
    ui = app.userInterface

    try:
        fileDialog = ui.createFileDialog()
        fileDialog.isMultiSelectEnabled = False
        fileDialog.title = (
            "Get the file to read from or the file to save the parameters to"
        )
        fileDialog.filter = "Text files (*.csv)"
        fileDialog.filterIndex = 0
        if isImport:
            dialogResult = fileDialog.showOpen()
        else:
            dialogResult = fileDialog.showSave()

        if dialogResult == adsk.core.DialogResults.DialogOK:  # type: ignore
            filename = fileDialog.filename
        else:
            return

        # if isImport is true read the parameters from a file
        if isImport:
            export(filename, do_stl, do_step, do_obj, do_3mf)
        else:
            write_parameters_to_file(filename)

    except:
        if ui:
            ui.messageBox("Failed:\n{}".format(traceback.format_exc()))


def write_parameters_to_file(filePath: str):
    app = adsk.core.Application.get()
    design: adsk.fusion.Design = app.activeProduct  # type: ignore

    with open(filePath, "w", newline="") as csvFile:
        csvWriter = csv.writer(csvFile, dialect=csv.excel)
        header = CSV_SPECIAL_HEADERS.copy()
        favorite_params = [param for param in design.userParameters if param.isFavorite]
        header += [param.name for param in favorite_params]
        csvWriter.writerow(header)
        expressions = ["default", "x"]
        expressions += [param.expression for param in favorite_params]
        csvWriter.writerow(expressions)

    # get the name of the file without the path
    partsOfFilePath = filePath.split("/")
    ui = app.userInterface
    ui.messageBox("Parameters written to " + partsOfFilePath[-1])


class ParameterList:
    def __init__(self, row: list[str], header: dict[int, str]):
        values = {name: row[column] for column, name in header.items()}
        self._export = bool(values[CSV_EXPORT_FLAG])
        self._export_name = values[CSV_EXPORT_NAME]
        self.params = {
            param: value
            for param, value in values.items()
            if param not in CSV_SPECIAL_HEADERS
        }

    @property
    def should_export(self):
        return self._export

    @property
    def output_filename(self):
        return self._export_name


def read_parameters_from_file(filePath: str):
    variations: list[ParameterList] = []
    with open(filePath) as csvFile:
        csvReader = csv.reader(csvFile, dialect=csv.excel)
        header: dict[int, str] = {}
        for line, row in enumerate(csvReader):
            if line == 0:
                header = {index: name for index, name in enumerate(row)}
            else:
                variations.append(ParameterList(row, header))
    return variations


def get_output_folder():
    app = adsk.core.Application.get()
    ui = app.userInterface
    folder_dialog = ui.createFolderDialog()
    folder_dialog.title = "Select output folder for exported models"
    result = folder_dialog.showDialog()
    if result != adsk.core.DialogResults.DialogOK:  # type: ignore
        return None
    return folder_dialog.folder


def export_meshes(
    output_folder: str,
    file_name: str,
    component: adsk.fusion.Component,
    do_stl: bool,
    do_step: bool,
    do_obj: bool,
    do_3mf: bool,
):
    export_manager = component.parentDesign.exportManager
    output_path = str(Path(output_folder) / file_name)
    if do_stl:
        futil.log("exporting stl")
        options = export_manager.createSTLExportOptions(component, output_path)
        export_manager.execute(options)
    if do_step:
        futil.log("exporting step")
        options = export_manager.createSTEPExportOptions(output_path, component)
        export_manager.execute(options)
    if do_obj:
        futil.log("exporting obj")
        options = export_manager.createOBJExportOptions(component, output_path)
        export_manager.execute(options)
    if do_3mf:
        futil.log("exporting 3mf")
        options = export_manager.createC3MFExportOptions(component, output_path)
        export_manager.execute(options)


def export(filePath: str, do_stl: bool, do_step: bool, do_obj: bool, do_3mf: bool):
    app = adsk.core.Application.get()
    design: adsk.fusion.Design = app.activeProduct  # type: ignore
    ui = app.userInterface
    output_folder = get_output_folder()
    if output_folder is None:
        return
    variations = read_parameters_from_file(filePath)
    for variation in variations:
        if variation.should_export:
            apply_parameters(ui, design, variation)
            export_meshes(
                output_folder,
                variation.output_filename,
                design.activeComponent,
                do_stl,
                do_step,
                do_obj,
                do_3mf,
            )
    ui.messageBox("Export finished successfully")


def apply_parameters(
    ui: adsk.core.UserInterface, design: adsk.fusion.Design, params: ParameterList
):
    try:
        paramsList = [oParam.name for oParam in design.allParameters]
        retryList: list[tuple[str, str]] = []

        for param_name, param_expression in params.params.items():
            if not update_parameter(
                ui, design, paramsList, param_name, param_expression
            ):
                retryList.append((param_name, param_expression))
        # let's keep going through the list until all is done
        count = 0
        while len(retryList) + 1 > count:
            count = count + 1
            flagged_successful: list[tuple[str, str]] = []
            for params_tuple in retryList:
                param_name, param_expression = params_tuple
                if update_parameter(
                    ui, design, paramsList, param_name, param_expression
                ):
                    flagged_successful.append(params_tuple)
            for resolved in flagged_successful:
                retryList.remove(resolved)

        if len(retryList) > 0:
            params_str = "\n".join([param for param, _ in retryList])

            ui.messageBox(f"Could not set the following parameters: {params_str}")
    except:
        if ui:
            ui.messageBox(f"AddIn Stop Failed:\n{traceback.format_exc()}")


def update_parameter(
    ui: adsk.core.UserInterface,
    design: adsk.fusion.Design,
    paramsList: list[str],
    param: str,
    expression: str,
):
    # get the values from the csv file.
    try:
        nameOfParam = param
        expressionOfParam = expression
    except Exception as e:
        print(str(e))
        # makes no sense to retry
        return True

    try:
        # if the name of the parameter is not an existing parameter let the user know
        if nameOfParam not in paramsList:
            ui.messageBox(
                f"A parameter with the name {nameOfParam} does not exist in the model. It will be ignored"
            )

        # update the values of existing parameters
        else:
            paramInModel = design.allParameters.itemByName(nameOfParam)
            paramInModel.expression = expressionOfParam
            print("Updated {}".format(nameOfParam))

        return True

    except Exception as e:
        print(str(e))
        print("Failed to update {}".format(nameOfParam))
        return False


def stop(_):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        obj_array: list[adsk.core.ToolbarControl | adsk.core.CommandDefinition] = []

        command_control_panel = command_control_by_id_for_panel(COMMAND_ID)
        if command_control_panel:
            obj_array.append(command_control_panel)

        command_definition = command_definition_by_id(COMMAND_ID)
        if command_definition:
            obj_array.append(command_definition)

        for obj in obj_array:
            destroy_object(ui, obj)

        futil.clear_handlers()
    except Exception:
        futil.handle_error("stop")
