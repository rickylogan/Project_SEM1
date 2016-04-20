<script type="text/javascript">
$(document).ready(function () {
            $('#btsubmit').click(function () {
                var txtUser = $('#txtName');
                var txtEmail = $('#txtEmail');
                if (txtUser.val().trim() == '') {
                    alert('Feild Name is not Empty');
                    return;
                }
                if (txtEmail.val().trim() == '') {
                    alert('Feild Email is not empty');
                }
                debugger;
                //else {
                //    var emailPattern = /^([a-zA-Z0-9_\-\.]+\@)([a-zA-Z0-9_\-\.])+\.([a-zA-Z]{2,4})$/;
                //    var code = emailPattern.exec(txtEmail.val());
                //    if(code == null)
                //    {
                //        alert('Your email is invalid! Please enter your email again');
                //    }
                //    else
                //    {
                //        document.getElementById('email').innerHTML = 'Email is invalid';
                //        document.getElementById('email').style.color = 'blue';
                //    }
                //}
            });
        });
        
        function hintName() {
            document.getElementById('name').innerHTML = 'Please enter name';
            document.getElementById('name').style.color = 'red';
        }

        function checkName() {
            var name = document.getElementById('txtName').value;
            if(name.length == 0)
            {
                document.getElementById('name').innerHTML = 'Feild name is not empty';
                document.getElementById('name').style.color = 'red';
            }
            else
            {
                document.getElementById('name').innerHTML = 'Name is invalid';
                document.getElementById('name').style.color = 'blue';
            }
        }        

        function hintEmail() {
            document.getElementById('email').innerHTML = 'Please enter email';
            document.getElementById('email').style.color = 'red';
        }

        function checkEmail() {
            var emailPattern = /^([a-zA-Z0-9_\-\.]+\@)([a-zA-Z0-9_\-\.])+\.([a-zA-Z]{2,4})$/;
            var code = emailPattern.exec(txtEmail.val());
            if (code == null) {
                document.getElementById('email').innerHTML = 'Your email is invalid! Please enter your email again';
                document.getElementById('email').style.color = 'red';
            }
            else {
                document.getElementById('email').innerHTML = 'Email is invalid';
                document.getElementById('email').style.color = 'blue';
            }
        }
		</script>