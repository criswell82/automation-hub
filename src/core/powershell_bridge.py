"""
PowerShell integration bridge for Automation Hub.
Provides seamless execution of PowerShell commands and scripts from Python.
"""

import subprocess
import logging
import os
from typing import Optional, Dict, Any, List
from pathlib import Path
import json


class PowerShellResult:
    """
    Represents the result of a PowerShell execution.
    """

    def __init__(
            self,
            stdout: str,
            stderr: str,
            return_code: int,
            command: str
    ):
        """
        Initialize a PowerShell result.

        Args:
            stdout: Standard output
            stderr: Standard error
            return_code: Exit code
            command: Executed command
        """
        self.stdout = stdout
        self.stderr = stderr
        self.return_code = return_code
        self.command = command
        self.success = return_code == 0

    def __str__(self) -> str:
        return f"PowerShellResult(success={self.success}, return_code={self.return_code})"

    def __repr__(self) -> str:
        return self.__str__()


class PowerShellBridge:
    """
    Bridge for executing PowerShell commands and scripts from Python.
    """

    def __init__(self, execution_policy: str = 'Bypass'):
        """
        Initialize the PowerShell bridge.

        Args:
            execution_policy: PowerShell execution policy to use
                            ('Bypass', 'RemoteSigned', 'Unrestricted', etc.)
        """
        self.logger = logging.getLogger(__name__)
        self.execution_policy = execution_policy
        self.powershell_path = self._find_powershell()

        if not self.powershell_path:
            self.logger.warning("PowerShell not found on system")

    def _find_powershell(self) -> Optional[str]:
        """
        Find PowerShell executable.

        Returns:
            str: Path to PowerShell executable or None
        """
        # Try PowerShell Core (pwsh) first, then Windows PowerShell
        for ps_exe in ['pwsh.exe', 'powershell.exe']:
            # Check if it's in PATH
            result = subprocess.run(
                ['where', ps_exe],
                capture_output=True,
                text=True,
                shell=True
            )

            if result.returncode == 0:
                path = result.stdout.strip().split('\n')[0]
                self.logger.info(f"Found PowerShell at: {path}")
                return path

        return None

    def execute_command(
            self,
            command: str,
            timeout: Optional[int] = None,
            capture_output: bool = True,
            as_json: bool = False
    ) -> PowerShellResult:
        """
        Execute a PowerShell command.

        Args:
            command: PowerShell command to execute
            timeout: Execution timeout in seconds
            capture_output: Whether to capture stdout/stderr
            as_json: If True, parse stdout as JSON

        Returns:
            PowerShellResult: Execution result
        """
        if not self.powershell_path:
            raise RuntimeError("PowerShell not available")

        # Build PowerShell command
        ps_command = [
            self.powershell_path,
            '-ExecutionPolicy', self.execution_policy,
            '-NoProfile',
            '-NonInteractive',
            '-Command', command
        ]

        self.logger.debug(f"Executing PowerShell command: {command[:100]}...")

        try:
            result = subprocess.run(
                ps_command,
                capture_output=capture_output,
                text=True,
                timeout=timeout,
                encoding='utf-8',
                errors='replace'
            )

            stdout = result.stdout if capture_output else ''
            stderr = result.stderr if capture_output else ''

            # Parse JSON if requested
            if as_json and stdout:
                try:
                    stdout = json.loads(stdout)
                except json.JSONDecodeError as e:
                    self.logger.warning(f"Failed to parse PowerShell output as JSON: {e}")

            ps_result = PowerShellResult(
                stdout=stdout,
                stderr=stderr,
                return_code=result.returncode,
                command=command
            )

            if not ps_result.success:
                self.logger.warning(f"PowerShell command failed: {stderr}")

            return ps_result

        except subprocess.TimeoutExpired:
            self.logger.error(f"PowerShell command timed out after {timeout} seconds")
            return PowerShellResult(
                stdout='',
                stderr=f'Command timed out after {timeout} seconds',
                return_code=-1,
                command=command
            )
        except Exception as e:
            self.logger.error(f"PowerShell execution failed: {e}")
            return PowerShellResult(
                stdout='',
                stderr=str(e),
                return_code=-1,
                command=command
            )

    def execute_script(
            self,
            script_path: str,
            parameters: Optional[Dict[str, Any]] = None,
            timeout: Optional[int] = None
    ) -> PowerShellResult:
        """
        Execute a PowerShell script file.

        Args:
            script_path: Path to PowerShell script (.ps1)
            parameters: Script parameters as dictionary
            timeout: Execution timeout in seconds

        Returns:
            PowerShellResult: Execution result
        """
        script_path = Path(script_path)

        if not script_path.exists():
            raise FileNotFoundError(f"PowerShell script not found: {script_path}")

        # Build command with parameters
        command = f"& '{script_path}'"

        if parameters:
            param_str = ' '.join([
                f"-{key} '{value}'" if isinstance(value, str) else f"-{key} {value}"
                for key, value in parameters.items()
            ])
            command += f" {param_str}"

        return self.execute_command(command, timeout=timeout)

    def execute_script_block(
            self,
            script: str,
            parameters: Optional[Dict[str, Any]] = None,
            timeout: Optional[int] = None,
            as_json: bool = False
    ) -> PowerShellResult:
        """
        Execute a multi-line PowerShell script block.

        Args:
            script: PowerShell script content
            parameters: Script parameters as dictionary
            timeout: Execution timeout in seconds
            as_json: If True, parse output as JSON

        Returns:
            PowerShellResult: Execution result
        """
        # Inject parameters into script
        if parameters:
            param_lines = []
            for key, value in parameters.items():
                if isinstance(value, str):
                    param_lines.append(f"${key} = '{value}'")
                elif isinstance(value, bool):
                    param_lines.append(f"${key} = ${str(value).lower()}")
                else:
                    param_lines.append(f"${key} = {value}")

            script = '\n'.join(param_lines) + '\n' + script

        return self.execute_command(script, timeout=timeout, as_json=as_json)

    def test_connection(self) -> bool:
        """
        Test if PowerShell is available and working.

        Returns:
            bool: True if PowerShell is available
        """
        if not self.powershell_path:
            return False

        result = self.execute_command("Write-Output 'test'")
        return result.success and 'test' in result.stdout

    def get_system_info(self) -> Dict[str, Any]:
        """
        Get system information using PowerShell.

        Returns:
            dict: System information
        """
        command = """
        $info = @{
            ComputerName = $env:COMPUTERNAME
            Username = $env:USERNAME
            OSVersion = [System.Environment]::OSVersion.Version.ToString()
            PSVersion = $PSVersionTable.PSVersion.ToString()
            Is64Bit = [Environment]::Is64BitOperatingSystem
        }
        $info | ConvertTo-Json
        """

        result = self.execute_command(command, as_json=True)

        if result.success and isinstance(result.stdout, dict):
            return result.stdout
        else:
            return {}

    def run_as_administrator(self, command: str) -> PowerShellResult:
        """
        Execute a PowerShell command with administrator privileges.
        Note: This will trigger UAC prompt.

        Args:
            command: Command to execute

        Returns:
            PowerShellResult: Execution result
        """
        admin_command = f"Start-Process powershell -Verb RunAs -ArgumentList '-ExecutionPolicy {self.execution_policy} -Command {command}' -Wait"

        return self.execute_command(admin_command)

    def get_environment_variable(self, name: str, default: Optional[str] = None) -> Optional[str]:
        """
        Get an environment variable using PowerShell.

        Args:
            name: Variable name
            default: Default value if not found

        Returns:
            str: Variable value or default
        """
        command = f"$env:{name}"
        result = self.execute_command(command)

        if result.success and result.stdout.strip():
            return result.stdout.strip()

        return default

    def set_environment_variable(
            self,
            name: str,
            value: str,
            scope: str = 'User'
    ) -> bool:
        """
        Set an environment variable using PowerShell.

        Args:
            name: Variable name
            value: Variable value
            scope: Variable scope ('User', 'Machine', or 'Process')

        Returns:
            bool: True if successful
        """
        command = f"[Environment]::SetEnvironmentVariable('{name}', '{value}', '{scope}')"
        result = self.execute_command(command)

        return result.success


# Singleton instance
_powershell_bridge: Optional[PowerShellBridge] = None


def get_powershell_bridge() -> PowerShellBridge:
    """Get the global PowerShellBridge instance."""
    global _powershell_bridge
    if _powershell_bridge is None:
        _powershell_bridge = PowerShellBridge()
    return _powershell_bridge
