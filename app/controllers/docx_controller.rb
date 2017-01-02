include DocxHelper

class DocxController < ApplicationController
  #unloadable

  def template_upload
    reset_template_names #updates names for version 1.1.0

    tracker = Tracker.find_by_name(params[:tracker])
    use_for_all = params[:use_for_all]
    uploaded_io = params[:template]
    if params[:template_action] == 'Add new template'
      append_template = true
    else
      append_template = false
      remove_templates_for(tracker)
    end

    upload_file(tracker, uploaded_io, append_template)

    if use_for_all == '1'
      source = list_templates_for(tracker).last
      Tracker.all.each do |t|
        unless t == tracker
          if append_template
            dest = 'files/export_docx/templates/' + t.name + (list_templates_for(t).count + 1).to_s + '.docx'
          else
            dest = 'files/export_docx/templates/' + t.name + '1.docx'
            remove_templates_for(t)
          end
          FileUtils.copy_file(source, dest)
        end
      end
    end

    redirect_to plugin_settings_path(Redmine::Plugin.find('export_docx'))
  end
  
  def template_download
    path_to_file = params[:path]
    if File.exist?(path_to_file)
      send_file(path_to_file)
    else
      flash[:error] = 'There is no template for ' + issue.tracker.name + ' issues.'
      redirect_to plugin_settings_path(Redmine::Plugin.find('export_docx'))
    end
  end
  
  def issue_export_docx
    issue = Issue.find(params[:id])
    if params[:file].present?
      template_path = 'files/export_docx/templates/' + params[:file]
    else
      template_path = list_templates_for(issue.tracker).last
    end
    issue_to_docx(issue, template_path)
    path_to_file = 'files/export_docx/export/' + issue.project.name + ' - ' + issue.tracker.name + ' #' + issue.id.to_s + '.docx'
    if File.exist?(path_to_file)
      send_file(path_to_file)
    end
  end
  
end