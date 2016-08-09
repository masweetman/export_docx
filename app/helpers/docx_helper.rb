require 'docx'

module DocxHelper

  def folder_structure
    FileUtils.mkdir_p 'files/export_docx/templates'
    FileUtils.mkdir_p 'files/export_docx/export'
  end

  def issue_to_docx(issue)
    FileUtils.rm_rf(Dir.glob('files/export_docx/export/*'))
    path_to_file = 'files/export_docx/templates/' + issue.tracker.name + '.docx'
    if File.exist?(path_to_file)
      doc = Docx::Document.open(path_to_file)
      doc.bookmarks.keys.each do |bookmark|
        # write standard issue fields
        case bookmark.downcase
        when 'project'
          doc.bookmarks[bookmark].insert_text_after(issue.project.name) unless issue.project.nil?
        when 'tracker'
          doc.bookmarks[bookmark].insert_text_after(issue.tracker.name) unless issue.tracker.nil?
        when 'id'
          doc.bookmarks[bookmark].insert_text_after(issue.id.to_s) unless issue.id.nil?
        when 'subject'
          doc.bookmarks[bookmark].insert_text_after(issue.subject) unless issue.subject.nil?
        when 'description'
          doc.bookmarks[bookmark].insert_multiple_lines(issue.description.lines.map(&:chomp)) unless issue.description.nil?
        when 'status'
          doc.bookmarks[bookmark].insert_text_after(issue.status.name) unless issue.status.nil?
        when 'priority'
          doc.bookmarks[bookmark].insert_text_after(issue.priority.name) unless issue.priority.nil?
        when 'author', 'added_by'
          doc.bookmarks[bookmark].insert_text_after(issue.author.name) unless issue.author.nil?
        when 'assignee', 'assigned_to'
          doc.bookmarks[bookmark].insert_text_after(issue.assigned_to.name) unless issue.assigned_to.nil?
        when 'category'
          doc.bookmarks[bookmark].insert_text_after(issue.category.name) unless issue.category.nil?
        when 'target_version', 'fixed_version'
          doc.bookmarks[bookmark].insert_text_after(issue.fixed_version.name) unless issue.fixed_version.nil?
        when 'start_date'
          doc.bookmarks[bookmark].insert_text_after(issue.start_date.strftime('%m/%d/%Y')) unless issue.start_date.nil?
        when 'due_date'
          doc.bookmarks[bookmark].insert_text_after(issue.due_date.strftime('%m/%d/%Y')) unless issue.due_date.nil?
        when 'created_on'
          doc.bookmarks[bookmark].insert_text_after(issue.created_on.strftime('%m/%d/%Y')) unless issue.created_on.nil?
        when 'closed_on'
          doc.bookmarks[bookmark].insert_text_after(issue.closed_on.strftime('%m/%d/%Y')) unless issue.closed_on.nil?
        when 'percent_done', 'done_ratio'
          doc.bookmarks[bookmark].insert_text_after(issue.done_ratio.to_s + '%') unless issue.done_ratio.nil?
        when 'estimated_time', 'estimated_hours'
          doc.bookmarks[bookmark].insert_text_after(issue.estimated_hours.to_s) unless issue.estimated_hours.nil?
        when 'spent_time', 'spent_hours'
          doc.bookmarks[bookmark].insert_text_after(issue.spent_hours.to_s) unless issue.spent_hours.nil?
        else
          #write custom issue fields
          custom_field = CustomField.find_by_name(bookmark.tr('_',' '))
          unless custom_field.nil?
            if custom_field.field_format == 'text'
              doc.bookmarks[bookmark].insert_multiple_lines(issue.custom_field_value(custom_field.id).lines.map(&:chomp)) unless issue.custom_field_value(custom_field.id).nil?
            elsif custom_field.field_format == 'list' && custom_field.multiple?
              doc.bookmarks[bookmark].insert_multiple_lines(issue.custom_field_value(custom_field.id)) unless issue.custom_field_value(custom_field.id).nil?
            elsif custom_field.field_format == 'user'
              doc.bookmarks[bookmark].insert_text_after(User.find(issue.custom_field_value(custom_field.id)).to_s) unless issue.custom_field_value(custom_field.id).nil?
            elsif custom_field.field_format == 'date'
              doc.bookmarks[bookmark].insert_text_after(issue.custom_field_value(custom_field.id).to_date.strftime('%m/%d/%Y')) unless issue.custom_field_value(custom_field.id).nil? || issue.custom_field_value(custom_field.id).empty?
            elsif custom_field.field_format == 'bool'
              if doc.bookmarks[bookmark].get_run_before.node.xpath('descendant::*').last.attributes['val'].nil?
                if issue.custom_field_value(custom_field.id) == '1'
                  doc.bookmarks[bookmark].insert_text_after('Yes')
                else
                  doc.bookmarks[bookmark].insert_text_after('No')
                end
              else
                doc.bookmarks[bookmark].get_run_before.node.xpath('descendant::*').last.attributes['val'].value = issue.custom_field_value(custom_field.id).to_s unless issue.custom_field_value(custom_field.id).nil?
              end
            else
              doc.bookmarks[bookmark].insert_text_after(issue.custom_field_value(custom_field.id).to_s)
            end
          end
        end
      end
      doc.save('files/export_docx/export/' + issue.project.name + ' - ' + issue.tracker.name + ' #' + issue.id.to_s + '.docx')
    else
      flash[:error] = 'There is no template for ' + issue.tracker.name + ' issues. Please notify your Redmine administrator.'
      redirect_to issue
    end
  end

end