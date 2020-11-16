require "csv"
require "java"
require "pathname"
java_import javax.swing.JOptionPane
java_import javax.swing.JFileChooser
java_import javax.swing.filechooser.FileNameExtensionFilter
java_import javax.swing.JDialog
java_import org.apache.commons.io.FileUtils
java_import javax.swing.UIManager

if UIManager.getLookAndFeel.getName == "Metal"
	UIManager.getInstalledLookAndFeels.each do |info|
		if info.getName == "Windows"
			UIManager.setLookAndFeel(info.getClassName)
			break
		end
	end
end

#Shows a dialog allowing the user to pick a choice
def show_options(message,options,default=nil,title="Options")
	if default.nil?
		default = options[0]
	end
	choice = JOptionPane.showInputDialog(nil,message,title,JOptionPane::PLAIN_MESSAGE,nil,options.to_java(:Object),default)
	return choice
end

#Shows a dialog allowing the user to select a directory
def prompt_directory(initial_directory=nil,title="Choose Directory")
	fc = JFileChooser.new
	fc.setDialogTitle(title)
	if !initial_directory.nil?
		fc.setCurrentDirectory(java.io.File.new(initial_directory))
	end
	fc.setFileSelectionMode(JFileChooser::DIRECTORIES_ONLY)
	if fc.showOpenDialog(nil) == JFileChooser::APPROVE_OPTION
		file = fc.getSelectedFile
		return file
	end
end

#Shows a dialog allowing the user to pick a file
def prompt_open_file(initial_directory=nil,filters=nil,title="Open File")
	fc = JFileChooser.new
	fc.setDialogTitle(title)
	if !filters.nil?
		fnef = nil
		filters.each do |k,v|
			fnef = FileNameExtensionFilter.new(k,*v)
			fc.addChoosableFileFilter(fnef)
		end
		fc.setFileFilter(fnef)
	end

	if !initial_directory.nil?
		fc.setCurrentDirectory(java.io.File.new(initial_directory))
	end

	if fc.showOpenDialog(nil) == JFileChooser::APPROVE_OPTION
		return fc.getSelectedFile
	else
		return java.io.File.new("")
	end
end

#Copies a directory, support paths as strings or java.io.File objects
def copy_directory(source,dest)
	if !source.is_a?(java.io.File)
		source = java.io.File.new(source)
	end

	if !dest.is_a?(java.io.File)
		dest = java.io.File.new(dest)
	end

	FileUtils.copyDirectory(source,dest)
end

#Locates cases in a directory by locating fbi2 files
#Returns array of case directories
def find_case_directories(search_directories)
	result = []
	search_directories = Array(search_directories)
	search_directories = search_directories.map{|d|!d.is_a?(String) ? d.getAbsolutePath : d }
	search_directories.each do |search_directory|
		fbi_paths = Dir.glob(File.join(search_directory, "**", "case.fbi2"))
		result += fbi_paths.map{|p|File.dirname(p)}
	end
	return result
end

#Timestamp to use on log and report files
time_stamp = Time.now.strftime("%Y%m%d_%H-%M-%S")

#======================#
# Simple logging class #
#======================#
class Logger
	class << self
		attr_accessor :log_file
		def log(obj)
			message = "#{Time.now.strftime("%Y%m%d %H:%M:%S")}: #{obj}"
			puts message
			File.open(@log_file,"a"){|f|f.puts message}
		end
	end
end
Logger.log_file = File.join(File.dirname(__FILE__), "#{time_stamp}_Log.txt")

#Prompt about making backups
message = "Do you want to make a backup of each case before migrating?"
backup_choices = ["Make Backups","No Backups Needed"]
backup_choice = show_options(message,backup_choices,nil,"Make Backups")
if backup_choice.nil?
	puts "User did not supply a backup choice"
	exit 1
end
make_backups = (backup_choice == backup_choices[0])

Logger.log "Make Backups: #{make_backups}"

#If they want to make backups, then we need to get the directory to store the backups
if make_backups
	backup_directory = prompt_directory("C:\\","Choose Backup Directory")
	if backup_directory.nil? || !backup_directory.exists?
		Logger.log "User did not provide a backup directory"
		exit 1
	end
	Logger.log "Backup Directory: #{backup_directory.getAbsolutePath}"
end

#Prompt user for how they want to supply the cases to be migrated
message = "How should the cases be specified?"
case_choices = ["Locate Cases in Directory","Provide Text File with Case Paths"]
case_choice = show_options(message,case_choices,nil,"Options")
if case_choice.nil?
	puts "User did not select a case input method"
	exit 1
end
locate_cases = (case_choice == case_choices[0])

Logger.log "Cases Input: #{case_choice}"

#Depending on case input choice we need to either ask them to supply a search directory
#or supply a text file containing case paths
case_paths = []
if locate_cases
	case_search_directory = prompt_directory("C:\\","Choose Case Search Directory")
	if case_search_directory.nil? || !case_search_directory.exists
		Logger.log "User did not specify a case search directory"
		exit 1
	else
		Logger.log "Case Search Directory: #{case_search_directory}"
		Logger.log "Locating cases (this may take a moment)..."
		case_paths = find_case_directories(case_search_directory.getAbsolutePath)
	end
else
	case_list_file = prompt_open_file(nil,{"Case Path List Text File (*.txt)"=>["txt"]},"Choose Case Paths Text File")

	if case_list_file.nil? || !case_list_file.exists || case_list_file.getAbsolutePath.empty?
		Logger.log "User did not provide an input list of case paths"
		exit 1
	else
		Logger.log "Case List File: #{case_list_file}"
		Logger.log "Loading case list file..."
		case_paths = File.readlines(case_list_file.getAbsolutePath,:encoding => "utf-8").map{|l|l.chomp}
	end
end

Logger.log "=== Cases to Migrate (#{case_paths.size}) ==="
case_paths.each do |path|
	Logger.log "\t#{path}"
end

#========================#
# Simple reporting class #
#========================#
class Reporter
	def initialize(report_file)
		@report_file = report_file
		CSV.open(@report_file,"w:utf-8"){|csv|csv << ["CasePath","Had Error","Error Message"]}
	end

	def report(case_path,had_error,message)
		CSV.open(@report_file,"a:utf-8"){|csv|csv << [case_path,had_error,message]}
	end

	def report_error(case_path,message)
		report(case_path,true,message)
	end

	def report_success(case_path)
		report(case_path,false,"")
	end
end

failures = []
successes = []

#Define report file and do some the work
report_file = File.join(File.dirname(__FILE__), "#{time_stamp}_Report.csv")
begin
	reporter = Reporter.new(report_file)

	case_paths.each_with_index do |case_path,index|
		Logger.log "===== (#{index+1}/#{case_paths.size}) : #{case_path} ====="
		#Make backups if requested, otherwise note that user opted out, catch exceptions
		#that may occur while doing this
		if make_backups
			begin
				Logger.log "(#{index+1}/#{case_paths.size}) Backing up case: #{case_path}"
				case_name_from_dir = Pathname.new(case_path).split.last.to_s
				case_backup_directory = File.join(backup_directory.getAbsolutePath, case_name_from_dir)
				Logger.log "(#{index+1}/#{case_paths.size}) Backup Destination: #{case_backup_directory}"
				java.io.File.new(case_backup_directory).mkdirs
				copy_directory(case_path,case_backup_directory)
			rescue Exception => exc
				message = "An error occurred while trying to create backup: "
				message << "Case: #{case_path}, "
				message << "Message: #{exc.message}"
				Logger.log message
				failures << case_path
				reporter.report_error(case_path,message)
				next
			end
		else
			Logger.log "(#{index+1}/#{case_paths.size}) User opted out of making backups!  Skipping backup of case."
		end

		#Migrate case (by opening it), catch exceptions that may occur during this
		begin
			Logger.log "(#{index+1}/#{case_paths.size}) Migrating case: #{case_path}"
			$current_case = $utilities.getCaseFactory.open(case_path,{"migrate"=>true})
			successes << case_path
			reporter.report_success(case_path)
		rescue Exception => exc
			message = "Error while trying to open/migrate case: #{case_path}, "
			message << exc.message
			Logger.log message
			failures << case_path
			reporter.report_error(case_path,message)
			next
		ensure
			if !$current_case.nil? && !$current_case.isClosed
				$current_case.close
			end
		end
	end

	#Dump a cursory count of successes and failures, point user to report CSV for more details
	Logger.log "\n\nSuccesses (#{successes.size}):"
	successes.each do |path|
		Logger.log "\t#{path}"
	end

	Logger.log "\nFailures (#{failures.size}):"
	failures.each do |path|
		Logger.log "\t#{path}"
	end
	if failures.size > 0
		Logger.log "See for error report: #{report_file}"
	end
rescue Exception => exc
	#Trap any exceptions that slipped by somehow
	Logger.log "Unexpected exception!"
	Logger.log "If you encounter this, send log file when reporting the issue!"
	Logger.log exc.message
	Logger.log exc.backtrace
end
