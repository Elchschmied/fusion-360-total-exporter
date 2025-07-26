"""
Modified TotalExport script for Autodesk Fusion 360.

This version enhances the original script by adding logic to avoid
re‑exporting designs that have already been saved.  When the script
discovers that a Fusion archive (.f3d or .f3z) for a design already
exists in the destination folder, it prompts the user to confirm if
the existing archive should be overwritten.  If the user declines,
the design is skipped and no new export is performed.  This helps
prevent unnecessary overwrites and allows selective updating of
existing exports.

The remainder of the script remains functionally identical to the
original: it walks all hubs, projects and folders accessible to the
current user, opens each design in the main thread and exports
Fusion archives, STEP files and DXF sketches.  Logging has been
retained so you can review what was exported and what was skipped.

Author: Justin Nesselrotte (original)
Modified by: ChatGPT
"""

from __future__ import with_statement

import adsk.core
import adsk.fusion
import adsk.cam
import traceback

from logging import Logger, FileHandler, Formatter
import logging
from threading import Thread

import time
import os
import re


class TotalExport(object):
    def __init__(self, app):
        self.app = app
        self.ui = self.app.userInterface
        self.data = self.app.data
        self.documents = self.app.documents
        self.log = Logger("Fusion 360 Total Export")
        # Ensure info messages (including project paths) are recorded in the log.
        self.log.setLevel(logging.INFO)
        self.num_issues = 0
        self.was_cancelled = False
        # Determines whether existing exported files should be overwritten.  This
        # flag is set once at the beginning of the export run and then
        # referenced for every subsequent file.  If True existing files are
        # always overwritten; if False they are always skipped.
        self.overwrite_existing: bool | None = None

        # Track progress of exported projects across runs.  This set holds tuples
        # of (hub_name, project_name) for projects that have been completely
        # processed.  The progress file path is set in run() after the output
        # directory is chosen.
        self.completed_projects = set()
        self.progress_path: str | None = None
        # Path to a log file listing all projects that were successfully exported.
        # This will be initialised in run() based on the selected output path.
        self.exported_projects_log_path: str | None = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def run(self, context):
        """Entry point for the export process."""
        self.ui.messageBox(
            "Searching for and exporting files will take a while, depending on how many files you have.\n\n"
            "You won't be able to do anything else. It has to do everything in the main thread and open and close every file.\n\n"
            "Take an early lunch."
        )

        output_path = self._ask_for_output_path()

        if output_path is None:
            return

        # Determine the progress file path and load any previously completed
        # projects.  The progress file records hub/project pairs that have
        # already been exported.  If the file exists, its contents populate
        # self.completed_projects; otherwise the set remains empty.
        self.progress_path = os.path.join(output_path, 'project_progress.tsv')
        # Determine the path for the exported projects log.  This file will
        # record which projects were successfully exported.
        self.exported_projects_log_path = os.path.join(output_path, 'exported_projects.log')
        self._load_progress()

        # If a progress file exists and contains entries, ask the user
        # whether to continue from the recorded progress or to start fresh.
        if os.path.exists(self.progress_path) and self.completed_projects:
            progress_prompt = (
                "Es existiert bereits eine Datei 'project_progress.tsv' mit bisher gesicherten Projekten.\n"
                "Möchten Sie den Export an dieser Stelle fortsetzen (Ja) oder von vorne beginnen (Nein)?"
            )
            progress_result = self.ui.messageBox(
                progress_prompt,
                "Fortschritt fortsetzen?",
                adsk.core.MessageBoxButtonTypes.YesNoButtonType
            )
            if progress_result == adsk.core.DialogResults.DialogNo:
                # Clear recorded progress and delete the progress file so the
                # export starts from scratch.  Also remove the exported
                # projects log file so it does not contain stale entries.
                self.completed_projects.clear()
                try:
                    os.remove(self.progress_path)
                except Exception:
                    pass
                # Remove the exported projects log if it exists
                try:
                    if self.exported_projects_log_path and os.path.exists(self.exported_projects_log_path):
                        os.remove(self.exported_projects_log_path)
                except Exception:
                    pass

        # Ask the user once whether existing exported files should be overwritten.
        # The result is stored in self.overwrite_existing and used for every file.
        overwrite_prompt = (
            "Es können bereits exportierte Dateien vorhanden sein.\n"
            "Möchten Sie vorhandene Dateien überschreiben?"
        )
        overwrite_result = self.ui.messageBox(
            overwrite_prompt,
            "Überschreiben?",
            adsk.core.MessageBoxButtonTypes.YesNoButtonType
        )
        # DialogYes (value 2) indicates the user clicked Yes, see Fusion API docs【457688972963656†screenshot】
        self.overwrite_existing = overwrite_result == adsk.core.DialogResults.DialogYes

        # Configure logging to write to a file in the output directory
        file_handler = FileHandler(os.path.join(output_path, 'output.log'))
        file_handler.setFormatter(Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        self.log.addHandler(file_handler)

        self.log.info("Starting export!")
        # Run the export in a loop so that if a connectivity error causes an
        # exception to propagate out of _export_data, the user has the option
        # to re-establish the connection and resume.  Without this loop,
        # unhandled exceptions would terminate the script.
        while True:
            try:
                self._export_data(output_path)
                break
            except BaseException as ex:
                # Log the exception and ask the user whether to retry the
                # entire export.  This typically covers cases where a
                # connectivity issue arises outside of the per-file handlers.
                self.log.exception("Exporting data failed", exc_info=ex)
                msg = (
                    f"Fehler beim Exportvorgang:\n{ex}\n\n"
                    "Möglicherweise ist die Verbindung zur Cloud unterbrochen.\n"
                    "Stellen Sie die Verbindung wieder her und klicken Sie auf 'Ja' zum erneuten Versuch.\n"
                    "Bei 'Nein' wird der Exportvorgang abgebrochen."
                )
                result = self.ui.messageBox(
                    msg,
                    "Verbindungsfehler",
                    adsk.core.MessageBoxButtonTypes.YesNoButtonType
                )
                if result == adsk.core.DialogResults.DialogYes:
                    time.sleep(5)
                    continue
                else:
                    # Stop exporting; leave loops
                    self.was_cancelled = True
                    break

        self.log.info("Done exporting!")

        if self.was_cancelled:
            self.ui.messageBox("Cancelled!")
        elif self.num_issues > 0:
            self.ui.messageBox(
                "The exporting process ran into {num_issues} issue{english_plurals}. "
                "Please check the log for more information".format(
                    num_issues=self.num_issues,
                    english_plurals="s" if self.num_issues > 1 else ""
                )
            )
        else:
            self.ui.messageBox("Export finished completely successfully!")

    def _export_data(self, output_path):
        """Iterate over all hubs and projects and export their contents."""
        progress_dialog = self.ui.createProgressDialog()
        progress_dialog.show("Exporting data!", "", 0, 1, 1)

        all_hubs = self.data.dataHubs
        for hub_index in range(all_hubs.count):
            hub = all_hubs.item(hub_index)

            self.log.info("Exporting hub \"{}\"".format(hub.name))

            all_projects = hub.dataProjects
            for project_index in range(all_projects.count):
                files = []
                project = all_projects.item(project_index)
                # Skip this project entirely if it has already been recorded as
                # completed in the progress file.  This allows the export to
                # resume after an interruption by continuing with the next
                # unprocessed project.
                if (hub.name, project.name) in self.completed_projects:
                    self.log.info(
                        "Skipping project \"{}\" – already recorded in progress file".format(project.name)
                    )
                    continue
                self.log.info("Exporting project \"{}\"".format(project.name))

                # Compute the base export directory for this project without creating it.
                # The structure is <output_path>/Hub <hubName>/Project <projectName>
                project_export_dir = os.path.join(
                    output_path,
                    "Hub {}".format(self._name(hub.name)),
                    "Project {}".format(self._name(project.name))
                )
                # Log and display the current backup path for the project.  The actual
                # directories are created later when exporting individual files.
                self.log.info(
                    "Sicherungspfad für Projekt \"{}\": {}".format(project.name, project_export_dir)
                )
                # We no longer display a message box for the project path because
                # requiring the user to click OK is disruptive.  The path is
                # recorded in the log instead.

                folder = project.rootFolder
                files.extend(self._get_files_for(folder))

                # Update the progress dialog message to include the current project name.  The
                # message now shows the hub and project indices as before, followed by the
                # project name.  The placeholders %v and %m will be replaced by the
                # progress dialog with the current file index and total count.
                progress_dialog.message = (
                    "Hub: {} of {}\n"
                    "Project: {} of {} - {}\n"
                    "Exporting design %v of %m"
                ).format(
                    hub_index + 1,
                    all_hubs.count,
                    project_index + 1,
                    all_projects.count,
                    project.name
                )
                progress_dialog.maximumValue = len(files)
                progress_dialog.reset()

                if not files:
                    self.log.info("No files to export for this project")
                    # Even if there are no files, mark the project as completed and
                    # record it in the exported projects log.  This ensures that
                    # empty projects are not repeatedly processed on subsequent runs.
                    self._append_progress(hub.name, project.name)
                    continue

                for file_index in range(len(files)):
                    if progress_dialog.wasCancelled:
                        self.log.info("The process was cancelled!")
                        self.was_cancelled = True
                        return

                    file = files[file_index]  # type: adsk.core.DataFile
                    # Update the progress dialog message to include the current file name.
                    # This provides more context to the user about which design is being
                    # exported.  Keep the placeholders %v and %m to show numeric progress.
                    progress_dialog.message = (
                        "Hub: {} of {}\n"
                        "Project: {} of {} - {}\n"
                        "Drawing: {}\n"
                        "Exporting design %v of %m"
                    ).format(
                        hub_index + 1,
                        all_hubs.count,
                        project_index + 1,
                        all_projects.count,
                        project.name,
                        file.name
                    )
                    progress_dialog.progressValue = file_index + 1
                    self._write_data_file(output_path, file)
                self.log.info("Finished exporting project \"{}\"".format(project.name))
                # Append this project to the progress file to mark it as completed.
                self._append_progress(hub.name, project.name)
            self.log.info("Finished exporting hub \"{}\"".format(hub.name))

    def _ask_for_output_path(self):
        """Prompt the user to select an output folder for the export."""
        folder_dialog = self.ui.createFolderDialog()
        folder_dialog.title = "Where should we store this export?"
        dialog_result = folder_dialog.showDialog()
        if dialog_result != adsk.core.DialogResults.DialogOK:
            return None
        output_path = folder_dialog.folder
        return output_path

    def _get_files_for(self, folder):
        """Recursively gather all DataFiles from the given folder."""
        files = []
        for file in folder.dataFiles:
            files.append(file)
        for sub_folder in folder.dataFolders:
            files.extend(self._get_files_for(sub_folder))
        return files

    def _write_data_file(self, root_folder, file):
        """Export a single DataFile and its contents.

        This function now checks if a Fusion archive has already been
        exported.  If an existing export is found, the user is asked
        whether to overwrite it; declining will skip the export.
        """
        # Only export Fusion designs (f3d/f3z). Skip others early
        if file.fileExtension not in ("f3d", "f3z"):
            self.log.info("Not exporting file \"{}\"".format(file.name))
            return

        self.log.info("Exporting file \"{}\"".format(file.name))

        # ---------------------------------------------------------------------
        # Early skip check: Determine the expected location of the exported
        # archive before opening the design.  If a local backup already
        # exists and the remote design hasn't changed (based on
        # dateModified), then skip exporting this file entirely.  This
        # prevents unnecessary calls to self.documents.open() for files
        # whose backups are already up to date.
        #
        # Build the relative path inside the project for this file by
        # walking up the folder hierarchy, just as is done later during
        # export.  Note: do not create any directories here — only compute
        # the path for comparison.
        file_folder = file.parentFolder
        relative_path = self._name(file_folder.name)
        tmp_folder = file_folder
        # Ascend the folder tree until no parentFolder is left
        while tmp_folder.parentFolder is not None:
            tmp_folder = tmp_folder.parentFolder
            relative_path = os.path.join(self._name(tmp_folder.name), relative_path)
        # The top-level folder's parentProject and parentHub identify where
        # this design belongs.
        parent_project_early = tmp_folder.parentProject
        parent_hub_early = parent_project_early.parentHub
        # Compute the full path to where the exported archive would reside.
        tentative_export_dir = os.path.join(
            root_folder,
            "Hub {}".format(self._name(parent_hub_early.name)),
            "Project {}".format(self._name(parent_project_early.name)),
            relative_path,
            self._name(file.name) + "." + file.fileExtension
        )
        tentative_export_base = os.path.join(tentative_export_dir, self._name(file.name))
        tentative_dest_archive = tentative_export_base + "." + file.fileExtension
        # If an archive exists and overwrite_existing is False (incremental backup),
        # compare timestamps and skip without opening the document if the
        # local backup is as recent or newer than the remote design.
        if os.path.exists(tentative_dest_archive) and not self.overwrite_existing:
            try:
                # Refresh metadata on the DataFile object.  Ignore errors.
                file.refresh()
            except BaseException:
                pass
            remote_ts = None
            # Try to handle both a raw epoch value or an adsk.core.DateTime
            try:
                remote_ts = float(file.dateModified)
            except Exception:
                try:
                    remote_date = file.dateModified
                    remote_ts = time.mktime((remote_date.year,
                                             remote_date.month,
                                             remote_date.day,
                                             remote_date.hour,
                                             remote_date.minute,
                                             remote_date.second,
                                             0, 0, -1))
                except Exception:
                    remote_ts = None
            if remote_ts is not None:
                try:
                    local_ts = os.path.getmtime(tentative_dest_archive)
                    if local_ts >= remote_ts:
                        self.log.info(
                            "Skipping file \"{}\" – local backup is up to date (early check).".format(file.name)
                        )
                        return
                except Exception:
                    # On failure to get local timestamp, proceed to open the file.
                    pass

        document = None
        # Attempt to open the document with a retry mechanism.  Network or
        # connectivity errors can cause this to fail; in that case the user
        # is prompted to retry or skip the file.
        while True:
            try:
                document = self.documents.open(file)
                if document is None:
                    raise Exception("Documents.open returned None")
                document.activate()
                break
            except BaseException as ex:
                # Log the exception and ask the user whether to retry opening
                # the file.  This typically covers connection drops or other
                # transient errors when accessing the cloud data.
                self.log.exception("Opening {} failed!".format(file.name), exc_info=ex)
                msg = (
                    f"Fehler beim Öffnen von '{file.name}':\n{ex}\n\n"
                    "Möglicherweise ist die Verbindung zur Cloud unterbrochen.\n"
                    "Stellen Sie die Verbindung wieder her und klicken Sie auf 'Ja' zum erneuten Versuch.\n"
                    "Bei 'Nein' wird dieser Export übersprungen."
                )
                result = self.ui.messageBox(
                    msg,
                    "Verbindungsfehler",
                    adsk.core.MessageBoxButtonTypes.YesNoButtonType
                )
                if result == adsk.core.DialogResults.DialogYes:
                    # Wait briefly before retrying
                    time.sleep(5)
                    continue
                else:
                    # User chose not to retry; skip this file
                    self.num_issues += 1
                    self.log.info(
                        "Skipping file '{}' due to open failure and user choice to skip.".format(file.name)
                    )
                    return

        try:
            # Build the folder structure used to store this export.  The
            # structure replicates the hierarchy in the data panel: Hub,
            # Project and folder names, followed by a folder named
            # <designName>.<extension> that will contain the exported
            # files.  This mirrors the behavior of the original script.
            file_folder = file.parentFolder
            file_folder_path = self._name(file_folder.name)
            # Walk up the folder tree to build the relative folder path
            while file_folder.parentFolder is not None:
                file_folder = file_folder.parentFolder
                file_folder_path = os.path.join(self._name(file_folder.name), file_folder_path)

            parent_project = file_folder.parentProject
            parent_hub = parent_project.parentHub

            # Construct (and create if necessary) the directory where
            # exports for this file will live
            export_dir = self._take(
                root_folder,
                "Hub {}".format(self._name(parent_hub.name)),
                "Project {}".format(self._name(parent_project.name)),
                file_folder_path,
                self._name(file.name) + "." + file.fileExtension
            )

            # Path to the Fusion archive that will be written (e.g. foo.f3d)
            file_export_base = os.path.join(export_dir, self._name(file.name))
            dest_archive = file_export_base + "." + file.fileExtension

            # A previous check was performed before opening the document to
            # determine whether this file should be skipped.  At this point,
            # either overwrite_existing is True or the remote design is newer,
            # so we unconditionally overwrite the existing archive (if any).

            self.log.info("Writing to \"{}\"".format(export_dir))

            fusion_document = adsk.fusion.FusionDocument.cast(document)
            design = fusion_document.design
            export_manager = design.exportManager

            # Export the Fusion archive (.f3d or .f3z).  If the execution fails
            # because of a connectivity issue, prompt the user to retry or
            # skip the export for this file.
            options = export_manager.createFusionArchiveExportOptions(file_export_base)
            while True:
                try:
                    export_manager.execute(options)
                    break
                except BaseException as ex:
                    self.log.exception(
                        "Fusion archive export failed for {}".format(file.name),
                        exc_info=ex
                    )
                    msg = (
                        f"Fehler beim Exportieren der Datei '{file.name}':\n{ex}\n\n"
                        "Möglicherweise ist die Verbindung zur Cloud unterbrochen.\n"
                        "Stellen Sie die Verbindung wieder her und klicken Sie auf 'Ja' zum erneuten Versuch.\n"
                        "Bei 'Nein' wird dieser Export übersprungen."
                    )
                    result = self.ui.messageBox(
                        msg,
                        "Verbindungsfehler",
                        adsk.core.MessageBoxButtonTypes.YesNoButtonType
                    )
                    if result == adsk.core.DialogResults.DialogYes:
                        time.sleep(5)
                        continue
                    else:
                        # Skip exporting this file on user request
                        self.num_issues += 1
                        self.log.info(
                            "Skipping file '{}' due to export error and user choice to skip.".format(file.name)
                        )
                        return

            # Export the root component and all of its sub-components
            self._write_component(export_dir, design.rootComponent)

            self.log.info("Finished exporting file \"{}\"".format(file.name))
        except BaseException as ex:
            self.num_issues += 1
            self.log.exception("Failed while working on \"{}\"".format(file.name), exc_info=ex)
            raise
        finally:
            # Close the document.  Suppress any errors to avoid
            # cascading failures.
            try:
                if document is not None:
                    document.close(False)
            except BaseException as ex:
                self.num_issues += 1
                self.log.exception("Failed to close \"{}\"".format(file.name), exc_info=ex)

    def _write_component(self, component_base_path, component):
        """Recursively export a component and its children into STEP and DXF files."""
        self.log.info("Writing component \"{}\" to \"{}\"".format(component.name, component_base_path))
        design = component.parentDesign

        output_path = os.path.join(component_base_path, self._name(component.name))

        # Export a STEP file for the component (skips if already exists)
        self._write_step(output_path, component)

        # Export DXF for all sketches in the component
        sketches = component.sketches
        for sketch_index in range(sketches.count):
            sketch = sketches.item(sketch_index)
            self._write_dxf(os.path.join(output_path, sketch.name), sketch)

        # Recursively export all child components
        occurrences = component.occurrences
        for occurrence_index in range(occurrences.count):
            occurrence = occurrences.item(occurrence_index)
            sub_component = occurrence.component
            sub_path = self._take(component_base_path, self._name(component.name))
            self._write_component(sub_path, sub_component)

    def _write_step(self, output_path, component):
        """Write a STEP file if it doesn't already exist."""
        file_path = output_path + ".stp"
        if os.path.exists(file_path):
            self.log.info("Step file \"{}\" already exists".format(file_path))
            return
        self.log.info("Writing step file \"{}\"".format(file_path))
        export_manager = component.parentDesign.exportManager
        options = export_manager.createSTEPExportOptions(output_path, component)
        export_manager.execute(options)

    def _write_stl(self, output_path, component):
        """Write an STL file if it doesn't already exist (not used by default)."""
        file_path = output_path + ".stl"
        if os.path.exists(file_path):
            self.log.info("Stl file \"{}\" already exists".format(file_path))
            return
        self.log.info("Writing stl file \"{}\"".format(file_path))
        export_manager = component.parentDesign.exportManager
        try:
            options = export_manager.createSTLExportOptions(component, output_path)
            export_manager.execute(options)
        except BaseException as ex:
            self.log.exception("Failed writing stl file \"{}\"".format(file_path), exc_info=ex)
            if component.occurrences.count + component.bRepBodies.count + component.meshBodies.count > 0:
                self.num_issues += 1

        bRepBodies = component.bRepBodies
        meshBodies = component.meshBodies
        if (bRepBodies.count + meshBodies.count) > 0:
            self._take(output_path)
            for index in range(bRepBodies.count):
                body = bRepBodies.item(index)
                self._write_stl_body(os.path.join(output_path, body.name), body)
            for index in range(meshBodies.count):
                body = meshBodies.item(index)
                self._write_stl_body(os.path.join(output_path, body.name), body)

    def _write_stl_body(self, output_path, body):
        """Write an STL file for a single body if it doesn't already exist."""
        file_path = output_path + ".stl"
        if os.path.exists(file_path):
            self.log.info("Stl body file \"{}\" already exists".format(file_path))
            return
        self.log.info("Writing stl body file \"{}\"".format(file_path))
        export_manager = body.parentComponent.parentDesign.exportManager
        try:
            options = export_manager.createSTLExportOptions(body, file_path)
            export_manager.execute(options)
        except BaseException:
            # Probably an empty model, ignore it
            pass

    def _write_iges(self, output_path, component):
        """Write an IGES file if it doesn't already exist (not used by default)."""
        file_path = output_path + ".igs"
        if os.path.exists(file_path):
            self.log.info("Iges file \"{}\" already exists".format(file_path))
            return
        self.log.info("Writing iges file \"{}\"".format(file_path))
        export_manager = component.parentDesign.exportManager
        options = export_manager.createIGESExportOptions(file_path, component)
        export_manager.execute(options)

    def _write_dxf(self, output_path, sketch):
        """Write a DXF file for a sketch if it doesn't already exist."""
        file_path = output_path + ".dxf"
        if os.path.exists(file_path):
            self.log.info("DXF sketch file \"{}\" already exists".format(file_path))
            return
        self.log.info("Writing dxf sketch file \"{}\"".format(file_path))
        sketch.saveAsDXF(file_path)

    def _take(self, *path):
        """Create a nested folder path and return its absolute path."""
        out_path = os.path.join(*path)
        os.makedirs(out_path, exist_ok=True)
        return out_path

    def _name(self, name):
        """Sanitize names by removing invalid filesystem characters.

        Also ensures that file names ending in .stp/.stl/.igs have
        underscores inserted before their extensions to avoid confusing
        directory names with file names.
        """
        name = re.sub('[^a-zA-Z0-9 \n\.]', '', name).strip()
        if name.endswith('.stp') or name.endswith('.stl') or name.endswith('.igs'):
            name = name[0: -4] + "_" + name[-3:]
        return name

    # -------------------------------------------------------------------------
    # Progress tracking helpers
    #
    # The export process can be interrupted, so it's useful to record which
    # projects have been successfully processed.  This table is stored in a
    # simple TSV file (hub\tproject per line) in the export root.  When the
    # script is restarted, it loads this file and skips any projects listed in
    # it.  After completing a project export, the hub/project name is appended
    # to the file.

    def _load_progress(self) -> None:
        """Load the progress file into self.completed_projects.

        If self.progress_path is defined and the file exists, read each line
        and add a (hub_name, project_name) tuple to the completed_projects set.
        This allows resuming an interrupted export by skipping previously
        exported projects.
        """
        # Ensure progress_path has been set
        if not self.progress_path:
            return
        self.completed_projects.clear()
        try:
            if os.path.exists(self.progress_path):
                with open(self.progress_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if not line:
                            continue
                        parts = line.split('\t')
                        if len(parts) >= 2:
                            hub_name, project_name = parts[0], parts[1]
                            self.completed_projects.add((hub_name, project_name))
        except Exception as ex:
            # Log but continue; progress tracking is optional
            self.log.exception("Failed to load progress file", exc_info=ex)

    def _append_progress(self, hub_name: str, project_name: str) -> None:
        """Append a completed project to the progress file.

        When a project has finished exporting, call this method to record its
        hub and project names.  It will create the file if necessary and
        append a new line using tab separation.  The (hub_name, project_name)
        tuple is also added to self.completed_projects.
        """
        if not self.progress_path:
            return
        try:
            # Append to the progress file (used for resuming the export).
            with open(self.progress_path, 'a', encoding='utf-8') as f:
                f.write(f"{hub_name}\t{project_name}\n")
            self.completed_projects.add((hub_name, project_name))
        except Exception as ex:
            # Log but don't interrupt the export
            self.log.exception("Failed to append to progress file", exc_info=ex)
        # Also append the project to the exported projects log.  Use a try/except
        # so that issues writing this auxiliary log do not interrupt the export.
        try:
            if self.exported_projects_log_path:
                with open(self.exported_projects_log_path, 'a', encoding='utf-8') as f_log:
                    f_log.write(f"{hub_name}\t{project_name}\n")
        except Exception as ex:
            self.log.exception("Failed to append to exported projects log", exc_info=ex)


def run(context):
    """Fusion 360 entry point.  Wraps the export class in a context manager."""
    ui = None
    try:
        app = adsk.core.Application.get()
        with TotalExport(app) as total_export:
            total_export.run(context)
    except:
        ui = app.userInterface
        ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
